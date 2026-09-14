"""
文件存储抽象层 - 支持本地存储和腾讯云COS
统一接口，通过配置切换存储后端
"""

import os
import io
import uuid
from datetime import datetime


class StorageBase:
    """存储基类"""

    def save(self, key, data):
        """保存文件，data可以是bytes或文件路径"""
        raise NotImplementedError

    def save_file(self, key, file_storage):
        """保存Flask上传的文件对象"""
        raise NotImplementedError

    def read(self, key):
        """读取文件内容，返回bytes"""
        raise NotImplementedError

    def download(self, key, local_path):
        """下载文件到本地路径"""
        raise NotImplementedError

    def delete(self, key):
        """删除文件"""
        raise NotImplementedError

    def exists(self, key):
        """判断文件是否存在"""
        raise NotImplementedError

    def get_url(self, key):
        """获取文件访问URL"""
        raise NotImplementedError

    def get_local_path(self, key):
        """获取本地路径（仅本地存储支持，COS返回None）"""
        return None


class LocalStorage(StorageBase):
    """本地文件存储"""

    def __init__(self, base_dir):
        self.base_dir = base_dir
        os.makedirs(self.base_dir, exist_ok=True)

    def _full_path(self, key):
        return os.path.join(self.base_dir, key.lstrip('/'))

    def save(self, key, data):
        path = self._full_path(key)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if isinstance(data, (bytes, bytearray)):
            with open(path, 'wb') as f:
                f.write(data)
        elif isinstance(data, str) and os.path.exists(data):
            # data是本地文件路径，复制
            import shutil
            shutil.copy2(data, path)
        return key

    def save_file(self, key, file_storage):
        path = self._full_path(key)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        file_storage.save(path)
        return path

    def read(self, key):
        path = self._full_path(key)
        with open(path, 'rb') as f:
            return f.read()

    def download(self, key, local_path):
        path = self._full_path(key)
        import shutil
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        shutil.copy2(path, local_path)
        return local_path

    def delete(self, key):
        path = self._full_path(key)
        if os.path.exists(path):
            os.remove(path)
            return True
        return False

    def exists(self, key):
        return os.path.exists(self._full_path(key))

    def get_url(self, key):
        # 本地存储返回相对路径
        return f'/files/{key}'

    def get_local_path(self, key):
        return self._full_path(key)


class TencentCOSStorage(StorageBase):
    """腾讯云对象存储 COS"""

    def __init__(self, secret_id, secret_key, region, bucket, domain=''):
        from qcloud_cos import CosConfig, CosS3Client
        self.secret_id = secret_id
        self.secret_key = secret_key
        self.region = region
        self.bucket = bucket
        self.domain = domain

        config = CosConfig(Region=region, SecretId=secret_id, SecretKey=secret_key)
        self.client = CosS3Client(config)

    def save(self, key, data):
        if isinstance(data, (bytes, bytearray)):
            self.client.put_object(
                Bucket=self.bucket,
                Body=data,
                Key=key,
            )
        elif isinstance(data, str) and os.path.exists(data):
            self.client.put_object_from_local_file(
                Bucket=self.bucket,
                LocalFilePath=data,
                Key=key,
            )
        return key

    def save_file(self, key, file_storage):
        # 先读到内存再上传
        data = file_storage.read()
        self.save(key, data)
        return key

    def read(self, key):
        response = self.client.get_object(Bucket=self.bucket, Key=key)
        return response['Body'].get_raw_stream().read()

    def download(self, key, local_path):
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        self.client.download_file(
            Bucket=self.bucket,
            Key=key,
            DestFilePath=local_path,
        )
        return local_path

    def delete(self, key):
        try:
            self.client.delete_object(Bucket=self.bucket, Key=key)
            return True
        except Exception:
            return False

    def exists(self, key):
        try:
            self.client.head_object(Bucket=self.bucket, Key=key)
            return True
        except Exception:
            return False

    def get_url(self, key):
        if self.domain:
            return f"https://{self.domain}/{key}"
        return f"https://{self.bucket}.cos.{self.region}.myqcloud.com/{key}"

    def get_local_path(self, key):
        # COS 不返回本地路径
        return None


class SupabaseStorage(StorageBase):
    """Supabase Storage 对象存储"""

    def __init__(self, supabase_url, supabase_key, bucket_name):
        try:
            from supabase import create_client
            self.client = create_client(supabase_url, supabase_key)
        except ImportError:
            raise ImportError("请安装 supabase: pip install supabase")
        self.bucket_name = bucket_name
        self.supabase_url = supabase_url

    def _ensure_bucket(self):
        """确保bucket存在，不存在则创建"""
        try:
            buckets = self.client.storage.list_buckets()
            names = [b.name for b in buckets]
            if self.bucket_name not in names:
                self.client.storage.create_bucket(self.bucket_name, public=True)
        except Exception:
            pass

    def save(self, key, data):
        self._ensure_bucket()
        if isinstance(data, (bytes, bytearray)):
            self.client.storage.from_(self.bucket_name).upload(key, data)
        elif isinstance(data, str) and os.path.exists(data):
            with open(data, 'rb') as f:
                self.client.storage.from_(self.bucket_name).upload(key, f.read())
        return key

    def save_file(self, key, file_storage):
        self._ensure_bucket()
        data = file_storage.read()
        self.save(key, data)
        return key

    def read(self, key):
        response = self.client.storage.from_(self.bucket_name).download(key)
        return response

    def download(self, key, local_path):
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        data = self.read(key)
        with open(local_path, 'wb') as f:
            f.write(data)
        return local_path

    def delete(self, key):
        try:
            self.client.storage.from_(self.bucket_name).remove([key])
            return True
        except Exception:
            return False

    def exists(self, key):
        try:
            # 列出目录中有没有这个文件
            folder = os.path.dirname(key)
            filename = os.path.basename(key)
            files = self.client.storage.from_(self.bucket_name).list(folder or None)
            return any(f['name'] == filename for f in files)
        except Exception:
            return False

    def get_url(self, key):
        # 生成公开访问URL（bucket需要设为public）
        res = self.client.storage.from_(self.bucket_name).get_public_url(key)
        return res.get('publicURL', '')


# ========== 全局存储实例 ==========
_upload_storage = None
_result_storage = None
_export_storage = None


def init_storage():
    """初始化存储实例（根据配置）"""
    global _upload_storage, _result_storage, _export_storage

    from config import STORAGE_TYPE, DATA_DIR

    if STORAGE_TYPE == 'cos':
        from config import COS_SECRET_ID, COS_SECRET_KEY, COS_REGION, COS_BUCKET, COS_DOMAIN
        _upload_storage = TencentCOSStorage(
            COS_SECRET_ID, COS_SECRET_KEY, COS_REGION, COS_BUCKET, COS_DOMAIN
        )
        _result_storage = TencentCOSStorage(
            COS_SECRET_ID, COS_SECRET_KEY, COS_REGION, COS_BUCKET, COS_DOMAIN
        )
        _export_storage = TencentCOSStorage(
            COS_SECRET_ID, COS_SECRET_KEY, COS_REGION, COS_BUCKET, COS_DOMAIN
        )
    elif STORAGE_TYPE == 'supabase':
        from config import SUPABASE_URL, SUPABASE_KEY, SUPABASE_BUCKET
        _upload_storage = SupabaseStorage(SUPABASE_URL, SUPABASE_KEY, SUPABASE_BUCKET)
        _result_storage = SupabaseStorage(SUPABASE_URL, SUPABASE_KEY, SUPABASE_BUCKET)
        _export_storage = SupabaseStorage(SUPABASE_URL, SUPABASE_KEY, SUPABASE_BUCKET)
    else:
        from config import UPLOAD_DIR, RESULT_DIR, EXPORT_DIR
        _upload_storage = LocalStorage(UPLOAD_DIR)
        _result_storage = LocalStorage(RESULT_DIR)
        _export_storage = LocalStorage(EXPORT_DIR)

    return _upload_storage, _result_storage, _export_storage


def get_upload_storage():
    """获取上传文件存储实例"""
    global _upload_storage
    if _upload_storage is None:
        init_storage()
    return _upload_storage


def get_result_storage():
    """获取结果文件存储实例"""
    global _result_storage
    if _result_storage is None:
        init_storage()
    return _result_storage


def get_export_storage():
    """获取导出文件存储实例"""
    global _export_storage
    if _export_storage is None:
        init_storage()
    return _export_storage
