"""Parse QTI assessment structure preserving section order and question pools."""

from dataclasses import dataclass
from typing import List, Optional, Union

from lxml import etree
from logzero import logger

QTI_NS = "{http://www.imsglobal.org/xsd/ims_qtiasiv1p2}"


@dataclass
class ItemBlock:
    xml_item: etree._Element


@dataclass
class GroupBlock:
    ident: str
    title: Optional[str]
    selection_count: Optional[int]
    items: List[etree._Element]


AssessmentBlock = Union[ItemBlock, GroupBlock]


def _local_tag(element: etree._Element) -> str:
    tag = element.tag
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _selection_count(section: etree._Element) -> Optional[int]:
    selection_number = section.find(
        f"./{QTI_NS}selection_ordering/{QTI_NS}selection/{QTI_NS}selection_number"
    )
    if selection_number is not None and selection_number.text:
        return int(selection_number.text)
    return None


def _parse_section_children(section: etree._Element) -> List[AssessmentBlock]:
    blocks: List[AssessmentBlock] = []

    for child in section:
        tag = _local_tag(child)
        if tag == "item":
            blocks.append(ItemBlock(xml_item=child))
        elif tag == "section":
            items = child.findall(f"./{QTI_NS}item")
            blocks.append(
                GroupBlock(
                    ident=child.get("ident", ""),
                    title=child.get("title"),
                    selection_count=_selection_count(child),
                    items=items,
                )
            )

    return blocks


def parse_structure(assessment_xml_path) -> List[AssessmentBlock]:
    """Return ordered assessment blocks from root_section."""
    try:
        root = etree.parse(str(assessment_xml_path)).getroot()
        root_section = root.find(f".//{QTI_NS}section[@ident='root_section']")
        if root_section is None:
            root_section = root.find(f".//{QTI_NS}section")
        if root_section is None:
            logger.error("No section found in assessment XML")
            return []
        return _parse_section_children(root_section)
    except OSError as e:
        logger.error("%s", e)
    except etree.ParseError as e:
        logger.error("XML parser error: %s", e)
    return []


def all_items_flat(assessment_xml_path) -> List[etree._Element]:
    """Return all items in document order (legacy flat export)."""
    try:
        root = etree.parse(str(assessment_xml_path)).getroot()
        return root.findall(f".//{QTI_NS}item")
    except OSError as e:
        logger.error("%s", e)
    except etree.ParseError as e:
        logger.error("XML parser error: %s", e)
    return []
