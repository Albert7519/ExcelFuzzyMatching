from django.db import models
import os


class ProcessedFile(models.Model):
    """存储处理过的Excel文件记录"""

    original_file = models.CharField(max_length=255)
    processed_file = models.CharField(max_length=255)
    upload_date = models.DateTimeField(auto_now_add=True)
    columns_processed = models.JSONField(default=list)
    # 新增字段记录处理模式和标准列
    processing_mode = models.CharField(
        max_length=20,
        choices=[("SELF_LEARNING", "自学习标准化"), ("REFERENCE", "参照标准匹配")],
        default="SELF_LEARNING",
        verbose_name="处理模式",
    )
    reference_column = models.CharField(
        max_length=255, null=True, blank=True, verbose_name="标准参照列"
    )

    def __str__(self):
        return os.path.basename(self.original_file)

    class Meta:
        verbose_name = "处理文件记录"
        verbose_name_plural = "处理文件记录"


class FuzzyMatchPattern(models.Model):
    """存储学习到的模糊匹配模式"""

    column_name = models.CharField(max_length=100, verbose_name="列名")
    original_pattern = models.CharField(max_length=255, verbose_name="原始模式")
    standardized_value = models.CharField(max_length=255, verbose_name="标准化值")

    class Meta:
        unique_together = ("column_name", "original_pattern")
        verbose_name = "模糊匹配模式"
        verbose_name_plural = "模糊匹配模式"

    def __str__(self):
        return (
            f"{self.column_name}: {self.original_pattern} -> {self.standardized_value}"
        )
