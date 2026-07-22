"""OneNote MCP Server — COM automation for OneNote desktop.

Provides 13 tools for navigating, reading, searching, and analyzing
OneNote content including embedded images and diagrams.
"""

import json
import xml.etree.ElementTree as ET

from mcp.server.fastmcp import FastMCP, Image

from onenote_lib import com_client, markdown_writer, write_gate
from onenote_lib.config import config
from onenote_lib.image_handler import get_all_images, get_image_base64
from onenote_lib.vision import describe_image, describe_images
from onenote_lib.xml_parser import (
    NotebookInfo,
    SectionGroupInfo,
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
    parse_section,
)

mcp = FastMCP(
    "OneNote MCP",
    instructions="Access OneNote desktop notebooks via COM automation. "
    "Read, search, and analyze pages including embedded images.",
)


# ── Navigation Tools ─────────────────────────────────────────────────


@mcp.tool()
def onenote_list_notebooks() -> str:
    """List all open notebooks with their IDs, names, paths, and last modified times."""
    xml = com_client.get_hierarchy("", com_client.NOTEBOOKS)
    notebooks = parse_notebooks(xml)
    result = []
    for nb in notebooks:
        result.append({
            "id": nb.id,
            "name": nb.name,
            "path": nb.path,
            "last_modified": nb.last_modified,
        })
    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_sections(notebook_id: str) -> str:
    """List all sections in a notebook, including sections inside section groups.

    Args:
        notebook_id: The notebook's OneNote ID (from onenote_list_notebooks)
    """
    xml = com_client.get_hierarchy(notebook_id, com_client.SECTIONS)
    notebooks = parse_notebooks(xml)
    if not notebooks:
        return json.dumps({"error": "Notebook not found"})

    nb = notebooks[0]
    result = _flatten_sections(nb.sections, nb.section_groups)
    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_pages(section_id: str) -> str:
    """List all pages in a section with titles and last modified times.

    Args:
        section_id: The section's OneNote ID (from onenote_list_sections)
    """
    xml = com_client.get_hierarchy(section_id, com_client.PAGES)
    section = parse_section(xml)
    if section is None:
        return json.dumps({"error": "Section not found"})

    return json.dumps(
        [
            {
                "id": p.id,
                "name": p.name,
                "last_modified": p.last_modified,
                "level": p.level,
            }
            for p in section.pages
        ],
        indent=2,
    )


@mcp.tool()
def onenote_get_notebook_tree(notebook_id: str = "", include_pages: bool = False) -> str:
    """Get the hierarchy: notebooks -> section groups -> sections (-> page titles).

    Page titles are omitted by default: across all notebooks they can run to
    hundreds of thousands of tokens. Scope to one notebook, or call
    onenote_list_pages for a single section, before asking for pages.

    Args:
        notebook_id: Optional notebook ID to scope the tree. Empty string = all notebooks.
        include_pages: Include page titles. Requires notebook_id to avoid huge output.
    """
    if include_pages and not notebook_id:
        return json.dumps({
            "error": "include_pages requires a notebook_id — the full page tree is too "
                     "large to return for all notebooks at once."
        })

    scope = com_client.PAGES if include_pages else com_client.SECTIONS
    xml = com_client.get_hierarchy(notebook_id, scope)
    notebooks = parse_notebooks(xml)
    return json.dumps([_notebook_to_tree(nb) for nb in notebooks], indent=2)


# ── Content Retrieval Tools ──────────────────────────────────────────


@mcp.tool()
def onenote_get_page(page_id: str) -> str:
    """Get a page's content as clean markdown. Images are listed as [Image N] references
    with callback IDs that can be retrieved with onenote_get_page_images or onenote_get_image.

    Args:
        page_id: The page's OneNote ID (from onenote_list_pages or search)
    """
    xml = com_client.get_page_content(page_id)
    markdown, images = parse_page_to_markdown(xml)

    if images:
        markdown += "\n\n---\n**Image References:**\n"
        for img in images:
            dims = ""
            if img.width and img.height:
                dims = f" ({img.width:.0f}x{img.height:.0f})"
            markdown += f"- [Image {img.index}]: callback_id=`{img.callback_id}`{dims}\n"

    return markdown


@mcp.tool()
def onenote_get_page_raw(page_id: str) -> str:
    """Get a page's raw OneNote XML content for debugging.

    Args:
        page_id: The page's OneNote ID
    """
    return com_client.get_page_content(page_id)


@mcp.tool()
def onenote_get_page_images(
    page_id: str,
    max_images: int = 10,
    max_size_kb: int = 512,
) -> list:
    """Extract all images from a page and return them as viewable images.
    Claude can see these images natively for analysis.

    Args:
        page_id: The page's OneNote ID
        max_images: Maximum number of images to extract (default 10)
        max_size_kb: Maximum size per image in KB (default 512, images are resized if larger)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)

    result = []
    for img in images:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
        else:
            result.append(f"[Image {img['index']}] (callback_id: {img['callback_id']})")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
def onenote_get_image(
    page_id: str,
    callback_id: str,
    max_size_kb: int = 512,
) -> list:
    """Get a single image by its callback ID. Returns the image for Claude to see natively.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID (from onenote_get_page output)
        max_size_kb: Maximum size in KB (default 512, image is resized if larger)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    return [Image(data=b64, media_type=media_type)]


# ── Search Tools ─────────────────────────────────────────────────────


@mcp.tool()
def onenote_search(query: str) -> str:
    """Full-text search across all open notebooks. Uses Windows Search indexing.

    Args:
        query: Search query string
    """
    xml = com_client.find_pages(query)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query})
    return json.dumps(results, indent=2)


@mcp.tool()
def onenote_search_in_notebook(notebook_id: str, query: str) -> str:
    """Search within a specific notebook.

    Args:
        notebook_id: The notebook's OneNote ID
        query: Search query string
    """
    xml = com_client.find_pages(query, notebook_id)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query, "notebook_id": notebook_id})
    return json.dumps(results, indent=2)


# ── Vision Analysis Tools ────────────────────────────────────────────


@mcp.tool()
async def onenote_analyze_page_visuals(
    page_id: str,
    prompt: str = "",
    max_images: int = 5,
    max_size_kb: int = 512,
) -> list:
    """Fetch all images from a page, send each to a vision model for description,
    and return both the descriptions and the raw images for Claude to see.

    Requires a vision-capable LLM server (set ONENOTE_VISION_URL and ONENOTE_VISION_MODEL).

    Args:
        page_id: The page's OneNote ID
        prompt: Optional custom prompt for the vision model
        max_images: Maximum number of images to process (default 5)
        max_size_kb: Maximum size per image in KB (default 512)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)
    analyzed = await describe_images(images, prompt or None)

    result = []
    for img in analyzed:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
            if "description" in img:
                result.append(f"Description: {img['description']}")
        else:
            result.append(f"[Image {img['index']}] Vision analysis: {img['description']}")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
async def onenote_describe_image(
    page_id: str,
    callback_id: str,
    prompt: str = "Describe this image in detail. If it's a diagram, explain the structure and relationships shown.",
    max_size_kb: int = 512,
) -> list:
    """Send a single image to the vision model with a custom prompt.
    Returns both the description and the raw image.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID
        prompt: Custom prompt for the vision model
        max_size_kb: Maximum size in KB (default 512)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    description = await describe_image(b64, media_type, prompt)

    return [
        f"Vision analysis: {description}",
        Image(data=b64, media_type=media_type),
    ]


# ── Write Tools (only registered when writing is enabled) ────────────


def _page_markdown(page_id: str) -> tuple[str, str, str]:
    """Return (page_xml, markdown, last_modified) for a page."""
    page_xml = com_client.get_page_content(page_id)
    markdown, _ = parse_page_to_markdown(page_xml)
    root = ET.fromstring(page_xml)
    return page_xml, markdown, root.get("lastModifiedTime", "")


def _sharing_note(page_id: str) -> str:
    """Describe whether the page lives in a notebook shared with other people."""
    try:
        object_id = page_id
        for _ in range(5):
            parent = com_client.get_hierarchy_parent(object_id)
            if not parent or parent == object_id:
                break
            object_id = parent
        notebooks = parse_notebooks(com_client.get_hierarchy("", com_client.NOTEBOOKS))
        for nb in notebooks:
            if nb.id == object_id:
                path = (nb.path or "").lower()
                if "sharepoint.com" in path:
                    return (
                        f"SHARED notebook {nb.name!r} (SharePoint). Edits sync to "
                        "everyone this notebook is shared with."
                    )
                return f"Personal notebook {nb.name!r}."
    except Exception:  # noqa: BLE001 - advisory only, never block the diff
        pass
    return "Could not determine which notebook this page belongs to."


def _register_write_tools() -> None:
    @mcp.tool()
    def onenote_create_page(section_id: str, title: str, markdown: str = "") -> str:
        """Create a new page in a section.

        Args:
            section_id: The section's OneNote ID where the page will be created
            title: Page title
            markdown: Optional page body as markdown. Headings, bullet and
                numbered lists, bold, italic and links are carried across.
        """
        try:
            new_page_id = com_client.create_new_page(section_id)
            page_xml = com_client.get_page_content(new_page_id)
            page_xml = markdown_writer.set_page_title(page_xml, title)
            if markdown:
                page_xml = markdown_writer.apply_markdown(page_xml, markdown, "append")
            com_client.update_page_content(page_xml)
            return json.dumps({"status": "created", "page_id": new_page_id, "title": title})
        except Exception as exc:  # noqa: BLE001 - surface to the model
            return json.dumps({"error": str(exc)})

    @mcp.tool()
    def onenote_diff_page(page_id: str, markdown: str, mode: str = "append") -> str:
        """Preview a page edit and get the confirm_token needed to apply it.

        Read-only. Show the returned diff to the user and get their agreement
        before calling onenote_update_page.

        Args:
            page_id: The page's OneNote ID
            markdown: The markdown to add, or to replace the page body with
            mode: "append" adds to the end of the page; "replace" discards the
                page's existing body, including images and ink
        """
        if mode not in ("append", "replace"):
            return json.dumps({"error": f"mode must be 'append' or 'replace', got {mode!r}"})

        page_xml, before, _ = _page_markdown(page_id)
        updated_xml = markdown_writer.apply_markdown(page_xml, markdown, mode)
        after, _ = parse_page_to_markdown(updated_xml)

        return json.dumps({
            "page_id": page_id,
            "mode": mode,
            "sharing": _sharing_note(page_id),
            "diff": write_gate.render_diff(before, after, page_id),
            "confirm_token": write_gate.issue_token(page_id, markdown, mode),
            "note": "Show this diff to the user. Only call onenote_update_page "
                    "once they have agreed to it.",
        }, indent=2)

    @mcp.tool()
    def onenote_update_page(
        page_id: str,
        markdown: str,
        confirm_token: str,
        mode: str = "append",
        force: bool = False,
    ) -> str:
        """Apply an edit previewed with onenote_diff_page.

        Args:
            page_id: The page's OneNote ID
            markdown: Must match exactly what was passed to onenote_diff_page
            confirm_token: The token returned by onenote_diff_page
            mode: Must match the mode passed to onenote_diff_page
            force: Apply even if the page changed since the diff, discarding
                the other edit. Leave false unless the user asks for it.
        """
        try:
            write_gate.consume_token(confirm_token, page_id, markdown, mode)
        except PermissionError as exc:
            return json.dumps({"error": str(exc)})

        page_xml, _, last_modified = _page_markdown(page_id)
        backup_path = write_gate.backup_page(page_id, page_xml)
        updated_xml = markdown_writer.apply_markdown(page_xml, markdown, mode)

        try:
            com_client.update_page_content(
                updated_xml,
                expected_last_modified=None if force else last_modified,
                force=force,
            )
        except com_client.PageConflictError as exc:
            return json.dumps({"error": str(exc), "backup": str(backup_path)})
        except Exception as exc:  # noqa: BLE001 - surface to the model, don't crash the server
            return json.dumps({"error": str(exc), "backup": str(backup_path)})

        return json.dumps({
            "status": "updated",
            "page_id": page_id,
            "mode": mode,
            "backup": str(backup_path),
        })


if config.enable_write:
    _register_write_tools()


# ── Helpers ──────────────────────────────────────────────────────────


def _flatten_sections(
    sections: list, section_groups: list, prefix: str = ""
) -> list[dict]:
    """Flatten sections and section groups into a flat list with group paths."""
    result = []
    for sec in sections:
        result.append({
            "id": sec.id,
            "name": sec.name,
            "group": prefix or None,
            "path": sec.path,
            "page_count": len(sec.pages),
        })
    for sg in section_groups:
        group_path = f"{prefix}/{sg.name}" if prefix else sg.name
        result.extend(_flatten_sections(sg.sections, sg.section_groups, group_path))
    return result


def _notebook_to_tree(nb: NotebookInfo) -> dict:
    """Convert a NotebookInfo to a tree dict."""
    tree = {
        "id": nb.id,
        "name": nb.name,
        "sections": [],
        "section_groups": [],
    }
    for sec in nb.sections:
        tree["sections"].append({
            "id": sec.id,
            "name": sec.name,
            "pages": [{"id": p.id, "name": p.name, "level": p.level} for p in sec.pages],
        })
    for sg in nb.section_groups:
        tree["section_groups"].append(_section_group_to_tree(sg))
    return tree


def _section_group_to_tree(sg: SectionGroupInfo) -> dict:
    """Convert a SectionGroupInfo to a tree dict."""
    tree = {
        "id": sg.id,
        "name": sg.name,
        "sections": [],
        "section_groups": [],
    }
    for sec in sg.sections:
        tree["sections"].append({
            "id": sec.id,
            "name": sec.name,
            "pages": [{"id": p.id, "name": p.name, "level": p.level} for p in sec.pages],
        })
    for child in sg.section_groups:
        tree["section_groups"].append(_section_group_to_tree(child))
    return tree


if __name__ == "__main__":
    mcp.run()
