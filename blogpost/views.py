from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse
from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
import openpyxl
from .models import Attendance
from .forms import AttendanceForm

class BlogList(ListView):
    model = Attendance
    template_name = 'list.html'

class BlogDetail(DetailView):
    model = Attendance
    template_name = 'detail.html'

class BlogCreate(CreateView):
    model = Attendance
    form_class = AttendanceForm
    template_name = 'create.html'
    success_url = reverse_lazy('list')

class BlogUpdate(UpdateView):
    model = Attendance
    form_class = AttendanceForm
    template_name = 'update.html'
    success_url = reverse_lazy('list')

class BlogDelete(DeleteView):
    model = Attendance
    template_name = 'delete.html'
    success_url = reverse_lazy('list')

def export_to_excel(request):
    data = Attendance.objects.all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出席データ"

    headers = ["授業名", "日付", "学籍番号", "出席・欠席・遅刻"]
    ws.append(headers)

    for item in data:
        ws.append([item.class_name, item.date, item.student_id, item.status])

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=attendance.xlsx'
    wb.save(response)
    return response
