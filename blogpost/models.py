from django.db import models

class Attendance(models.Model):
    class_name = models.CharField(max_length=100)
    date = models.DateField()
    student_id = models.CharField(max_length=20)
    status = models.CharField(max_length=10, choices=[('出席', '出席'), ('欠席', '欠席'), ('遅刻', '遅刻')])
    image = models.ImageField(upload_to='attendance_images/', blank=True, null=True)  # 画像フィールドを追加

    def __str__(self):
        return f"{self.class_name} - {self.student_id} - {self.status}"

