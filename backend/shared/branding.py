"""AdaptSCM product identifiers (API paths, S3, file/source domains)."""

API_PREFIX = "/api/adaptscm"

FILE_DOMAIN_STOCK = "adaptscm"
FILE_DOMAIN_STOCK_LEGACY = "inventory"
FILE_DOMAIN_SHIPMENT = "shipment"

SOURCE_STOCK = "ADAPTSCM"
SOURCE_STOCK_LEGACY = "INVENTORY"
SOURCE_SHIPMENT = "SHIPMENT"

S3_PREFIX_DEFAULT = "adaptscm"


def normalize_file_domain(file_domain: str | None) -> str:
    fd = (file_domain or FILE_DOMAIN_STOCK).strip().lower()
    if fd in (FILE_DOMAIN_STOCK, FILE_DOMAIN_STOCK_LEGACY):
        return FILE_DOMAIN_STOCK
    return fd


def is_stock_file_domain(file_domain: str | None) -> bool:
    fd = (file_domain or FILE_DOMAIN_STOCK).strip().lower()
    return fd in (FILE_DOMAIN_STOCK, FILE_DOMAIN_STOCK_LEGACY)


def s3_prefix_from_settings(settings) -> str:
    return (getattr(settings, "s3_prefix", None) or S3_PREFIX_DEFAULT).strip().strip("/")
