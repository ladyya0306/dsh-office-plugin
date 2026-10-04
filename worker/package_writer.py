"""Repackage Office bytes while replacing only existing authorized parts."""
import io
import zipfile


def replace_parts(original: bytes, replacements: dict[str, bytes]) -> bytes:
    with zipfile.ZipFile(io.BytesIO(original)) as archive:
        infos = archive.infolist()
        names = {info.filename for info in infos}
        if len(names) != len(infos):
            raise ValueError("ZIP部件名称重复")
        if any(name not in names or name.endswith("/") for name in replacements):
            raise ValueError("只允许替换已存在的文件部件")
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as destination:
            destination.comment = archive.comment
            for info in infos:
                destination.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
        return output.getvalue()
