from django.db import models


class RollFile(models.Model):
    title = models.TextField(verbose_name="파일명", default="학교")
    roll_file = models.FileField(upload_to="rolls/%Y/")

    def __str__(self):
        return self.title

    class Meta:
        verbose_name = "명렬표 파일"
        verbose_name_plural = "명렬표 파일"
