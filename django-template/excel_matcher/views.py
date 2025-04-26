import os
import json
import tempfile
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
    """处理Excel文件上传（不再本地保存，仅存内存）"""
    if request.method == "POST" and request.FILES.get("file"):
        excel_file = request.FILES["file"]

        # 检查文件扩展名
        if not excel_file.name.endswith((".xlsx", ".xls")):
            return JsonResponse({"error": "请上传Excel文件(.xlsx或.xls)"})

        try:
            # 读取文件内容到内存
            file_bytes = excel_file.read()
            filename = excel_file.name
            # 用临时文件获取列名（Windows下需delete=False，手动删除）
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=os.path.splitext(filename)[1]
            ) as tmp:
                tmp.write(file_bytes)
                tmp.flush()
                temp_path = tmp.name
            try:
                service = ExcelService()
                columns = service.get_excel_columns(temp_path)
            finally:
                os.remove(temp_path)
            # 存储文件内容和文件名到session
            request.session["uploaded_file_bytes"] = (
                file_bytes.hex()
            )  # 存为16进制字符串，兼容session
            request.session["uploaded_file_name"] = filename
            # 清除旧的本地路径
            if "uploaded_file_path" in request.session:
                del request.session["uploaded_file_path"]
            return JsonResponse(
                {
                    "success": True,
                    "message": "文件上传成功",
                    "columns": columns,
                    "filename": filename,
                }
            )
        except Exception as e:
            return JsonResponse({"error": str(e)})
    return JsonResponse({"error": "未找到上传的文件"})


def _ensure_uploaded_file_on_disk(request):
    """如果上传文件未落盘，则写入临时文件并返回路径"""
    file_path = request.session.get("uploaded_file_path")
    if file_path and os.path.exists(file_path):
        return file_path
    file_hex = request.session.get("uploaded_file_bytes")
    filename = request.session.get("uploaded_file_name")
    if file_hex and filename:
        file_bytes = bytes.fromhex(file_hex)
        # 只在 processed 目录下生成临时文件
        processed_dir = os.path.join(settings.MEDIA_ROOT, "processed")
        os.makedirs(processed_dir, exist_ok=True)
        temp_path = os.path.join(processed_dir, f"temp_{filename}")
        with open(temp_path, "wb") as f:
            f.write(file_bytes)
        # 存回session，后续可直接用
        request.session["uploaded_file_path"] = temp_path
        return temp_path
    return None


@csrf_exempt
def preview_matching(request):
    """预览Excel文件的匹配结果"""
    if request.method == "POST":
        try:
            data = json.loads(request.body)
            columns_to_match = data.get("columns_to_match", [])
            threshold = int(data.get("threshold", 80))
            processing_mode = data.get("processing_mode", "SELF_LEARNING")
            reference_column = data.get("reference_column")
            # 获取本地文件路径（如无则写入临时文件）
            file_path = _ensure_uploaded_file_on_disk(request)
            if not file_path or not os.path.exists(file_path):
                return JsonResponse({"error": "找不到上传的文件，请重新上传"})
            if not columns_to_match:
                return JsonResponse({"error": "请选择至少一个需要匹配的列"})
            if processing_mode == "REFERENCE" and (
                not reference_column or reference_column not in columns_to_match
            ):
                return JsonResponse(
                    {"error": "参照标准匹配模式需要选择一个有效的标准列"}
                )
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
            processing_mode = data.get("processing_mode", "SELF_LEARNING")
            reference_column = data.get("reference_column")
            # 获取本地文件路径（如无则写入临时文件）
            file_path = _ensure_uploaded_file_on_disk(request)
            if not file_path or not os.path.exists(file_path):
                return JsonResponse({"error": "找不到上传的文件，请重新上传"})
            if not columns_to_match:
                return JsonResponse({"error": "请选择至少一个需要匹配的列"})
            if processing_mode == "REFERENCE" and (
                not reference_column or reference_column not in columns_to_match
            ):
                return JsonResponse(
                    {"error": "参照标准匹配模式需要选择一个有效的标准列"}
                )
            service = ExcelService()
            processed_file_path = service.process_excel_file(
                file_path,
                columns_to_match,
                threshold,
                processing_mode,
                reference_column,
            )
            ProcessedFile.objects.create(
                original_file=file_path,
                processed_file=processed_file_path,
                columns_processed=columns_to_match,
                processing_mode=processing_mode,
                reference_column=reference_column,
            )
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
    """下载处理后的Excel文件，文件名为源文件名+_processed，下载后立即删除文件"""
    file_path = request.session.get("processed_file_path")
    if not file_path or not os.path.exists(file_path):
        return HttpResponse("找不到处理后的文件，请重新处理", status=404)
    original_filename = request.session.get("uploaded_file_name", "result.xlsx")
    name, ext = os.path.splitext(original_filename)
    download_filename = f"{name}_processed{ext}"
    # 先读取文件内容到内存
    with open(file_path, "rb") as f:
        file_data = f.read()
    # 删除本地文件
    try:
        os.remove(file_path)
    except Exception:
        pass
    response = HttpResponse(
        file_data,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{download_filename}"'
    return response
