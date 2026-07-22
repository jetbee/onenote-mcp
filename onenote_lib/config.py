import os
from pathlib import Path

from pydantic_settings import BaseSettings


class OneNoteConfig(BaseSettings):
    model_config = {"env_prefix": "ONENOTE_"}

    vision_url: str = "http://localhost:1234"
    vision_model: str = ""
    vision_fallback_url: str = ""
    vision_fallback_model: str = ""
    max_image_size_kb: int = 512
    max_images_per_page: int = 20

    # Writing is opt-in. A OneNote notebook shared through SharePoint or
    # OneDrive syncs to everyone it is shared with, so an edit made here is not
    # a local change. Set ONENOTE_ENABLE_WRITE=1 to expose the write tools.
    enable_write: bool = False
    # Where the pre-edit copy of a page is kept. Defaults to a per-user data
    # directory; these files contain the full plain-text content of your pages.
    backup_dir: str = ""

    def resolved_backup_dir(self) -> Path:
        if self.backup_dir:
            return Path(self.backup_dir)
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "onenote-mcp" / "backups"


config = OneNoteConfig()
