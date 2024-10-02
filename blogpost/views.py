from django.shortcuts import render, redirect
import openpyxl
from .models import Attendance
import os
from django.conf import settings
from datetime import datetime
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login,logout
from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from .forms import AttendanceForm
from django.db import IntegrityError

# 出席一覧の表示
class BlogList(ListView):
    model = Attendance
    template_name = 'list.html'

# 出席詳細の表示
class BlogDetail(DetailView):
    model = Attendance
    template_name = 'detail.html'

# 出席の新規作成
class BlogCreate(CreateView):
    model = Attendance
    form_class = AttendanceForm
    template_name = 'create.html'
    success_url = reverse_lazy('list')

# 出席情報の更新
class BlogUpdate(UpdateView):
    model = Attendance
    form_class = AttendanceForm
    template_name = 'update.html'
    success_url = reverse_lazy('list')

# 出席情報の削除
class BlogDelete(DeleteView):
    model = Attendance
    template_name = 'delete.html'
    success_url = reverse_lazy('list')

# 出席情報をエクセルにエクスポートし、画像ファイルのリンクを保存
def export_to_excel(request):
    data = Attendance.objects.all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出席データ"

    # ヘッダーを追加
    headers = ["授業名", "日付", "学籍番号", "出席・欠席・遅刻", "画像ファイルリンク"]
    ws.append(headers)

    for idx, item in enumerate(data, start=2):
        # 出席情報を追加
        ws.append([item.class_name, item.date, item.student_id, item.status])

        # 画像のURLをセルに保存
        if item.image:
            # 画像のURLを生成 (MEDIA_URLを使ってWebアクセス可能なURLにする)
            img_url = f"{request.build_absolute_uri(settings.MEDIA_URL)}{str(item.image)}"
            print(f"画像のURL: {img_url}")  # デバッグ用にURLを出力

            # エクセルのセルにリンクを追加
            link = f'=HYPERLINK("{img_url}", "画像を表示")'
            ws.cell(row=idx, column=5).value = link  # 画像ファイルリンクをE列に挿入

    # ファイル名を生成 (日時を含めると便利)
    filename = f"attendance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    file_path = os.path.join(settings.MEDIA_ROOT, filename)

    # エクセルファイルをサーバーに保存
    try:
        wb.save(file_path)
        print(f"エクセルファイルが保存されました: {file_path}")  # 保存先を出力
    except Exception as e:
        print(f"エクセルファイルの保存中にエラーが発生しました: {str(e)}")

    # 保存後にリダイレクト
    return redirect('list')  # 適切なビューにリダイレクト

# サインアップ（新規ユーザー登録）
def signupview(request):
    if request.method == 'POST':
        username_data = request.POST['username_data']
        password_data = request.POST['password_data']
        try:
            # ユーザーを作成
            User.objects.create_user(username=username_data, password=password_data)
            return render(request, 'signup.html', {'success': 'ユーザー登録が完了しました。'})
        except IntegrityError:
            # ユーザーが既に存在する場合、エラーメッセージを表示
            return render(request, 'signup.html', {'error': 'このユーザーは既に登録されています。'})
    else:
        return render(request, 'signup.html', {})

# ログインビュー
def loginview(request):
    if request.method == 'POST':
        username_data = request.POST['username_data']
        password_data = request.POST['password_data']
        user = authenticate(request, username=username_data, password=password_data)
        if user is not None:
            login(request, user)
            return redirect('list')
        else:
            return render(request, 'login.html', {'error': 'ログインに失敗しました。'})
    return render(request, 'login.html')

# サンプルビュー（必要に応じて使用）
def sampleview(request):
    if request.method == 'POST':
        return redirect('login')
    else:
        return render(request, 'login.html', {})
    
def logoutview(request):
    logout(request)
    return redirect('login')

