from django.db import models

class Eleve(models.Model):
    SERIE_CHOICES = [
        ('SN', 'العلوم الطبيعية'),
        ('M', 'الرياضيات'),
        ('LM', 'الآداب العصرية'),
        ('LO', 'الآداب الأصلية'),
        ('TM', 'الشعبة الفنية'),
        ('TS','الرياضيات و الهندسة الكهربائية'),
        ('LA', 'اللغات'),
    ]
        
    STATUT_CHOICES = [
        ('Admis', 'ناجح'),
        ('Ajourné', 'راسب'),
        ('Sessionaire', 'دورة ثانية'),
        ('Absent', 'غائب'),
    ]

    num_table = models.CharField(max_length=20, unique=True, verbose_name="رقم المقعد")
    nom_complet = models.CharField(max_length=255, verbose_name="الاسم الكامل")
    serie = models.CharField(max_length=10, choices=SERIE_CHOICES, verbose_name="الشعبة")
    wilaya = models.CharField(max_length=100, verbose_name="الولاية")
    centre = models.CharField(max_length=150, verbose_name="المركز")
    etablissement = models.CharField(max_length=200, verbose_name="المدرسة / الثانوية")
    moyenne = models.FloatField(verbose_name="المعدل")
    statut = models.CharField(max_length=20, choices=STATUT_CHOICES, verbose_name="النتيجة")

    def __str__(self):
        return f"{self.num_table} - {self.nom_complet}"