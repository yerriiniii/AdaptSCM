import boto3

from shared.config import get_runtime_settings


def is_s3_configured() -> bool:
    settings = get_runtime_settings()
    return settings.s3_enabled


def get_s3_client():
    settings = get_runtime_settings()
    if not settings.s3_enabled:
        raise RuntimeError("S3 환경변수가 아직 모두 설정되지 않았습니다.")

    return boto3.client(
        "s3",
        aws_access_key_id=settings.aws_access_key_id,
        aws_secret_access_key=settings.aws_secret_access_key,
        region_name=settings.aws_region,
    )


def create_presigned_upload_url(*, bucket: str, key: str, content_type: str | None = None, expires_in: int = 3600) -> str:
    client = get_s3_client()
    params = {"Bucket": bucket, "Key": key}
    if content_type:
        params["ContentType"] = content_type
    return client.generate_presigned_url(
        "put_object",
        Params=params,
        ExpiresIn=expires_in,
        HttpMethod="PUT",
    )
