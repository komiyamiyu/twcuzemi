# admin.py

from django.contrib import admin
from .models import Attendance  # 正しいモデル名をインポート

admin.site.register(Attendance)

