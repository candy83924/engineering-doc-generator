"""ZIP packaging service for bundling generated documents."""

import io
import zipfile


class PackagingService:
    """Package multiple generated files into a ZIP archive."""

    def package(self, files: dict[str, bytes]) -> bytes:
        """
        Package generated files into a ZIP.

        Args:
            files: {filename: file_bytes} mapping

        Returns:
            ZIP archive as bytes
        """
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, content in files.items():
                zf.writestr(filename, content)
        buffer.seek(0)
        return buffer.getvalue()

    def package_with_metadata(
        self, files: dict[str, bytes], metadata_text: str | None = None
    ) -> bytes:
        """Package files with an optional metadata/readme file."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, content in files.items():
                zf.writestr(filename, content)
            if metadata_text:
                zf.writestr("_生成資訊.txt", metadata_text.encode("utf-8"))
        buffer.seek(0)
        return buffer.getvalue()
