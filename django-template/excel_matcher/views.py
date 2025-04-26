import os
import json
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from .models import ProcessedFile
from .services.excel_service import ExcelService


def index(request):
    """首页视图，显示文件上传表单"""
    return render(request, "excel_matcher/index.html")


@csrf_exempt
def upload_file(request):
    """处理Excel文件上传"""
    if request.method == "POST" and request.FILES.get("file"):
        excel_file = request.FILES["file"]

        # 检查文件扩展名
        if not excel_file.name.endswith((".xlsx", ".xls")):
            return JsonResponse({"error": "请上传Excel文件(.xlsx或.xls)"})

        try:
            # 保存上传的文件
            service = ExcelService()
            file_path = service.save_uploaded_file(excel_file)

            # 获取Excel文件的列名
            columns = service.get_excel_columns(file_path)

            # 存储文件路径到session，以便后续处理
            request.session["uploaded_file_path"] = file_path

            return JsonResponse(
                {
                    "success": True,
                    "message": "文件上传成功",
                    "columns": columns,
                    "filename": os.path.basename(file_path),
                }
            )

        except Exception as e:
            return JsonResponse({"error": str(e)})

    return JsonResponse({"error": "未找到上传的文件"})


@csrf_exempt
def preview_matching(request):
    """预览Excel文件的匹配结果"""
    if request.method == "POST":
        try:
            data = json.loads(request.body)
            columns_to_match = data.get("columns_to_match", [])
            threshold = int(data.get("threshold", 80))

            # 获取处理模式和标准列
            processing_mode = data.get("processing_mode", "SELF_LEARNING")
            reference_column = data.get("reference_column")

            # 从session获取上传的文件路径
            file_path = request.session.get("uploaded_file_path")

            if not file_path or not os.path.exists(file_path):
                return JsonResponse({"error": "找不到上传的文件，请重新上传"})

            if not columns_to_match:
                return JsonResponse({"error": "请选择至少一个需要匹配的列"})

            # 参照标准模式需要有有效的标准列
            if processing_mode == "REFERENCE" and (
                not reference_column or reference_column not in columns_to_match
            ):
                return JsonResponse(
                    {"error": "参照标准匹配模式需要选择一个有效的标准列"}
                )

            # 预览处理结果
            service = ExcelService()
            preview_results = service.preview_matches(
                file_path,
                columns_to_match,
                threshold,
                processing_mode,
                reference_column,
            )

            return JsonResponse(
                {
                    "success": True,
                    "message": "预览生成成功",
                    "preview_results": preview_results,
                }
            )

        except Exception as e:
            return JsonResponse({"error": f"生成预览时出错: {str(e)}"})

    return JsonResponse({"error": "无效的请求方法"})


@csrf_exempt
def process_file(request):
    """处理Excel文件的模糊匹配"""
    if request.method == "POST":
        try:
            data = json.loads(request.body)
            columns_to_match = data.get("columns_to_match", [])
            threshold = int(data.get("threshold", 80))

            # 获取处理模式和标准列
            processing_mode = data.get("processing_mode", "SELF_LEARNING")
            reference_column = data.get("reference_column")

            # 从session获取上传的文件路径
            file_path = request.session.get("uploaded_file_path")

            if not file_path or not os.path.exists(file_path):
                return JsonResponse({"error": "找不到上传的文件，请重新上传"})

            if not columns_to_match:
                return JsonResponse({"error": "请选择至少一个需要匹配的列"})

            # 参照标准模式需要有有效的标准列
            if processing_mode == "REFERENCE" and (
                not reference_column or reference_column not in columns_to_match
            ):
                return JsonResponse(
                    {"error": "参照标准匹配模式需要选择一个有效的标准列"}
                )

            # 处理文件
            service = ExcelService()
            processed_file_path = service.process_excel_file(
                file_path,
                columns_to_match,
                threshold,
                processing_mode,
                reference_column,
            )

            # 记录处理记录
            ProcessedFile.objects.create(
                original_file=file_path,
                processed_file=processed_file_path,
                columns_processed=columns_to_match,
                processing_mode=processing_mode,
                reference_column=reference_column,
            )

            # 存储处理后的文件路径到session
            request.session["processed_file_path"] = processed_file_path

            return JsonResponse(
                {
                    "success": True,
                    "message": "文件处理成功",
                    "processed_file": os.path.basename(processed_file_path),
                }
            )

        except Exception as e:
            return JsonResponse({"error": f"处理文件时出错: {str(e)}"})

    return JsonResponse({"error": "无效的请求方法"})


def download_file(request):
    """下载处理后的Excel文件"""
    # 从session获取处理后的文件路径
    file_path = request.session.get("processed_file_path")

    if not file_path or not os.path.exists(file_path):
        return HttpResponse("找不到处理后的文件，请重新处理", status=404)

    # 提供文件下载
    filename = os.path.basename(file_path)
    response = FileResponse(open(file_path, "rb"))
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response
