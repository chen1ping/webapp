"""生成自签名SSL证书，支持多种访问方式"""
import datetime
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

subject = issuer = x509.Name([
    x509.NameAttribute(NameOID.COUNTRY_NAME, "CN"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "ChenPingAn"),
    x509.NameAttribute(NameOID.COMMON_NAME, "chenpinganyyds"),
])

cert = (
    x509.CertificateBuilder()
    .subject_name(subject)
    .issuer_name(issuer)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(datetime.datetime.now(datetime.timezone.utc))
    .not_valid_after(datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=3650))
    .add_extension(
        x509.SubjectAlternativeName([
            x509.DNSName("chenpinganyyds"),
            x509.DNSName("localhost"),
            x509.DNSName("*.local"),
            x509.IPAddress(__import__("ipaddress").IPv4Address("127.0.0.1")),
            x509.IPAddress(__import__("ipaddress").IPv4Address("192.168.1.107")),
            x509.IPAddress(__import__("ipaddress").IPv4Address("192.168.1.0")),
        ]),
        critical=False,
    )
    .add_extension(
        x509.ExtendedKeyUsage([x509.ExtendedKeyUsageOID.SERVER_AUTH]),
        critical=False,
    )
    .sign(key, hashes.SHA256())
)

with open("C:\\WebApp\\cert.pem", "wb") as f:
    f.write(cert.public_bytes(serialization.Encoding.PEM))

with open("C:\\WebApp\\key.pem", "wb") as f:
    f.write(key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    ))

print("证书生成成功: cert.pem, key.pem")
print(f"有效期: 10年")
print(f"支持域名: chenpinganyyds, localhost")
print(f"支持IP: 127.0.0.1, 192.168.1.107")
