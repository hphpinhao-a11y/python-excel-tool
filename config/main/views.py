from django.shortcuts import render
from django.http import HttpResponse
from .report import clean_dataset, dataframe_to_excel_response
import pandas as pd
import os
from django.conf import settings
from io import BytesIO


def upload_excel(request):

    # =========================
    # STEP 1：上傳檔案
    # =========================
    if request.method == 'POST' and request.FILES.get('excel_file'):

        file = request.FILES['excel_file']

        # 建立 uploads 資料夾
        upload_dir = os.path.join(settings.MEDIA_ROOT, "uploads")
        os.makedirs(upload_dir, exist_ok=True)

        file_path = os.path.join(upload_dir, file.name)

        # 寫入檔案到硬碟
        with open(file_path, 'wb+') as destination:
            for chunk in file.chunks():
                destination.write(chunk)

        # 只讀 Excel 結構（不讀資料）
        excel = pd.ExcelFile(file_path)
        sheet_names = excel.sheet_names

        # session 只存「檔案路徑」
        request.session['uploaded_file_path'] = file_path

        return render(request, 'select_sheet.html', {
            'sheet_names': sheet_names
        })


    # =========================
    # STEP 2：使用者選擇 sheet
    # =========================
    elif request.method == 'POST' and request.POST.getlist('sheets'):

        selected_sheets = request.POST.getlist('sheets')
        file_path = request.session.get('uploaded_file_path')

        # 防呆：session過期或檔案不存在
        if not file_path or not os.path.exists(file_path):
            return render(request, 'upload.html', {
                'error': '檔案已失效，請重新上傳'
            })

        # 讀檔成 bytes（避免 Windows 鎖檔）
        with open(file_path, "rb") as f:
            file_bytes = f.read()

        excel_buffer = BytesIO(file_bytes)

        # ⭐ 保留 N.D. / n.a. / 空白 的關鍵設定
        sheets = pd.read_excel(
            excel_buffer,
            sheet_name=selected_sheets,
            dtype=str,
            keep_default_na=False
        )

        all_cleaned = []

        # 每個 sheet 清洗
        for name, sheet_df in sheets.items():
            cleaned = clean_dataset(sheet_df)
            cleaned['SourceSheet'] = name
            all_cleaned.append(cleaned)

        # 合併全部 sheet
        cleaned_df = pd.concat(all_cleaned, ignore_index=True)

        # 轉為 Excel bytes
        excel_file = dataframe_to_excel_response(cleaned_df)

        response = HttpResponse(
            excel_file,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="cleaned.xlsx"'

        # ⭐ 非常重要：刪除暫存檔（解決 WinError 32）
        try:
            os.remove(file_path)
        except:
            pass

        # 清除 session
        if 'uploaded_file_path' in request.session:
            del request.session['uploaded_file_path']

        return response

    # =========================
    # 初始頁面
    # =========================
    return render(request, 'upload.html')








