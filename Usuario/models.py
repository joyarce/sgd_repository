from django.db import models
from django.utils import timezone

class FilePreview(models.Model):
    blob_name = models.CharField(max_length=255, unique=True)
    signed_url = models.TextField(null=True, blank=True)
    expires_at = models.DateTimeField(null=True, blank=True)

    def is_expired(self):
        return timezone.now() >= self.expires_at

