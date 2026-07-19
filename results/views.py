import os
import io
import openpyxl
import tempfile  # <-- ضروري جداً
from django.shortcuts import render, redirect
from django.contrib import messages
from .models import Eleve
# 1. الصفحة الرئيسية: البحث أولاً ثم الأوائل الثلاثة
def home_view(request):
    series = ['SN', 'M', 'LO', 'LM', 'TM', 'TS', 'LA']
    top_students = {}
    for s in series:
        top_students[s] = Eleve.objects.filter(serie=s, statut__icontains='admis').order_by('-moyenne')[:3]

    student_result = None
    rank_in_serie = None
    searched = False

    if 'q' in request.GET:
        searched = True
        query = request.GET.get('q', '').strip()
        if query:
            try:
                student_result = Eleve.objects.get(num_table=query)
                status_val = str(student_result.statut or '').lower()
                
                # حساب الترتيب فقط إذا كان ناجحاً
                if 'admis' in status_val or 'ناجح' in status_val:
                    higher_count = Eleve.objects.filter(
                        serie=student_result.serie,
                        statut__icontains='admis',
                        moyenne__gt=student_result.moyenne
                    ).count()
                    rank_in_serie = higher_count + 1
            except Eleve.DoesNotExist:
                student_result = None

    context = {
        'top_students': top_students,
        'student': student_result,
        'rank': rank_in_serie,
        'searched': searched  # هذا المتغير سيتحكم في إخفاء كروت المتفوقين
    }
    return render(request, 'home.html', context)


# 2. صفحة الفرز والإحصائيات الذكية والمتجاوبة
# 2. صفحة الفرز والإحصائيات الذكية والمتجاوبة (نسخة الفلترة الديناميكية تلقائياً)
def stats_view(request):
    # 1. استقبال قيم الفلاتر من الطلب (GET)
    selected_wilaya = request.GET.get('wilaya', '')
    selected_serie = request.GET.get('serie', '')
    selected_school = request.GET.get('etablissement', '')
    selected_status = request.GET.get('status', '')

    # 2. بناء قوائم الخيارات بشكل ديناميكي من قاعدة البيانات لكافة الفلاتر
    wilayas = Eleve.objects.values_list('wilaya', flat=True).distinct().order_by('wilaya')
    series_list = Eleve.objects.values_list('serie', flat=True).distinct().order_by('serie')
    status_list = Eleve.objects.values_list('statut', flat=True).distinct().order_by('statut')

    # 3. التصفية الذكية للمؤسسات حسب الولاية والشعبة لتسهيل الاختيار التلقائي
    school_filter = Eleve.objects.all()
    if selected_wilaya:
        school_filter = school_filter.filter(wilaya=selected_wilaya)
    if selected_serie:
        school_filter = school_filter.filter(serie=selected_serie)
        
    schools = school_filter.values_list('etablissement', flat=True).distinct().order_by('etablissement')

    # 4. فلترة مشروطة: لا يتم جلب أي بيانات (None) إلا بعد اختيار الولاية كشرط للبدء
    students = Eleve.objects.none()
    has_filtered = False

    # تفعيل البحث والنتائج فقط في حال قام المستخدم باختيار ولاية محددة
    if selected_wilaya:
        has_filtered = True
        students = Eleve.objects.filter(wilaya=selected_wilaya)
        
        if selected_serie:
            students = students.filter(serie=selected_serie)
        if selected_school:
            students = students.filter(etablissement=selected_school)
            
        if selected_status:
            if selected_status in ['Admis', 'ناجح']:
                students = students.filter(statut__icontains='admis')
            elif selected_status in ['Ajourné', 'راسب']:
                students = students.filter(statut__icontains='ajourn')
            elif selected_status in ['Sessionaire', 'دورة ثانية', 'دورة تكميلية']:
                students = students.filter(statut__icontains='sess')
            elif selected_status in ['Absent', 'غائب']:
                students = students.filter(statut__icontains='abs')
            else:
                students = students.filter(statut=selected_status)

    total_count = students.count() if has_filtered else 0

    context = {
        'wilayas': wilayas,
        'series_list': series_list,
        'schools': schools,
        'status_list': status_list,
        'students': students,
        'total_count': total_count,
        'has_filtered': has_filtered,
        'selected_wilaya': selected_wilaya,
        'selected_serie': selected_serie,
        'selected_school': selected_school,
        'selected_status': selected_status,
    }
    return render(request, 'stats.html', context)

# 3. رفع ملف الإكسيل وقراءة العناوين
def upload_excel_view(request):
    if request.method == "POST" and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        
        # حفظ الملف في مجلد مؤقت آمن بالسيرفر بدلاً من الـ RAM
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    for chunk in excel_file.chunks():
        temp_file.write(chunk)
        temp_file.close()
        
        # حفظ مسار الملف فقط في الـ Session
    request.session['temp_excel_path'] = temp_file.name
        
    try:
            workbook = openpyxl.load_workbook(temp_file.name, read_only=True)
            sheet = workbook.active
            headers = [str(cell.value).strip() for cell in next(sheet.iter_rows(max_row=1))]
            workbook.close()
            return render(request, 'mapping.html', {'headers': headers})
    except Exception as e:
            messages.error(request, f"خطأ في قراءة الملف: {e}")
            return render(request, 'upload.html')
            
    return render(request, 'upload.html')

def import_mapped_data_view(request):
    temp_path = request.session.get('temp_excel_path')
    if not temp_path or not os.path.exists(temp_path):
        messages.error(request, "انتهت الجلسة، يرجى إعادة الرفع.")
        return redirect('upload_excel')
        
    try:
        # قراءة الملف من القرص الصلب مباشرة (سريع جداً)
        workbook = openpyxl.load_workbook(temp_path, read_only=True, data_only=True)
        sheet = workbook.active
        
        # ... (بقية منطقك في استخراج الأعمدة كما هو) ...
        # (تأكد من ترك منطق الـ idx_table وغيرها كما كانت)
        
        Eleve.objects.all().delete()
        students_batch = []
        total_saved = 0
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row: continue
            # ... (باقي منطق معالجة الطلاب الذي كتبته أنت) ...
            
            # ... (كود الـ bulk_create الخاص بك) ...
            
        workbook.close()
        
        # تنظيف: حذف الملف المؤقت من السيرفر
        if os.path.exists(temp_path):
            os.remove(temp_path)
        del request.session['temp_excel_path']
        
        messages.success(request, f"تم بنجاح حفظ {total_saved} طالب.")
    except Exception as e:
        messages.error(request, f"خطأ: {e}")
        
    return redirect('upload_excel')