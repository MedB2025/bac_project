import os
import openpyxl
from django.conf import settings
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
        # إضافة الترتيب التنازلي حسب المعدل مباشرة عند بداية الفلترة
        students = Eleve.objects.filter(wilaya=selected_wilaya).order_by('-moyenne')
        
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
# 3. رفع ملف الإكسيل وقراءة العناوين
def upload_excel_view(request):
    if request.method == "POST" and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        
        temp_dir = os.path.join(settings.BASE_DIR, 'temp_excel')
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        file_path = os.path.join(temp_dir, 'uploaded_data.xlsx')
        
        with open(file_path, 'wb+') as destination:
            for chunk in excel_file.chunks():
                destination.write(chunk)
                
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            sheet = workbook.active
            headers = [str(cell.value).strip() for cell in next(sheet.iter_rows(max_row=1))]
            workbook.close()
            
            return render(request, 'mapping.html', {'headers': headers})
        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)
            messages.error(request, f"خطأ في قراءة ملف الإكسيل: {e}")
            return render(request, 'upload.html')
            
    return render(request, 'upload.html')

# 4. معالجة واستيراد البيانات المطابقة مع دعم كافة الحالات (ناجح، راسب، تكميلية، غائب)
def import_mapped_data_view(request):
    if request.method == "POST":
        file_path = os.path.join(settings.BASE_DIR, 'temp_excel', 'uploaded_data.xlsx')
        if not os.path.exists(file_path):
            messages.error(request, "لم يتم العثور على الملف.")
            return redirect('upload_excel')
            
        try:
            idx_table = int(request.POST.get('col_table'))
            idx_nom = int(request.POST.get('col_nom'))
            idx_wilaya = int(request.POST.get('col_wilaya'))
            idx_etablissement = int(request.POST.get('col_etablissement'))
            idx_centre = int(request.POST.get('col_centre'))
            idx_serie = int(request.POST.get('col_serie'))
            idx_moyenne = int(request.POST.get('col_moyenne'))
            idx_statut = int(request.POST.get('col_statut'))
        except:
            messages.error(request, "تأكد من اختيار كافة الأعمدة.")
            return redirect('upload_excel')

        students_to_create = []

        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet = workbook.active
            
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row: continue
                required_len = max(idx_table, idx_nom, idx_wilaya, idx_etablissement, idx_centre, idx_serie, idx_moyenne, idx_statut) + 1
                if len(row) < required_len: continue
                
                num_table = str(row[idx_table]).strip()
                if not num_table or num_table.lower() == "none": continue
                
                moyenne_raw = str(row[idx_moyenne]).replace(',', '.').strip()
                try: moyenne_val = float(moyenne_raw)
                except: moyenne_val = 0.0
                
                # الذكاء المطور لمعالجة كافة الحالات بدون أي أخطاء نصية
                raw_statut = str(row[idx_statut]).strip().lower()
                
                if 'adm' in raw_statut or 'ناجح' in raw_statut:
                    final_statut = 'Admis'
                elif 'sess' in raw_statut or 'ثاني' in raw_statut or 'تكميل' in raw_statut:
                    final_statut = 'Sessionaire'
                elif 'abs' in raw_statut or 'غائب' in raw_statut or 'غايب' in raw_statut:
                    final_statut = 'Absent'
                elif 'ajou' in raw_statut or 'راسب' in raw_statut:
                    final_statut = 'Ajourné'
                else:
                    # صمام أمان مبني على المعدل في حال عدم تطابق النص نهائياً
                    if moyenne_val >= 10.0:
                        final_statut = 'Admis'
                    else:
                        final_statut = 'Ajourné'

                # صمام أمان إضافي: حتى لو قرأ النظام من ملف الإكسيل نصاً خاطئاً وكان المعدل ناجحاً، يتم تصحيحه فوراً
                if moyenne_val >= 10.0 and final_statut == 'Ajourné':
                    final_statut = 'Admis'

                student = Eleve(
                    num_table=num_table,
                    nom_complet=str(row[idx_nom]).strip(),
                    wilaya=str(row[idx_wilaya]).strip(),
                    etablissement=str(row[idx_etablissement]).strip(),
                    centre=str(row[idx_centre]).strip(),
                    serie=str(row[idx_serie]).strip().upper(),
                    moyenne=moyenne_val,
                    statut=final_statut
                )
                students_to_create.append(student)

            workbook.close()

            if students_to_create:
                Eleve.objects.all().delete()
                Eleve.objects.bulk_create(students_to_create, batch_size=500)
                messages.success(request, f"تمت المزامنة بنجاح وحفظ {len(students_to_create)} طالب بكافة حالاتهم المختلفة.")
        except Exception as e:
            messages.error(request, f"حدث خطأ: {e}")
        finally:
            if os.path.exists(file_path): os.remove(file_path)

        return redirect('upload_excel')
    return redirect('upload_excel')