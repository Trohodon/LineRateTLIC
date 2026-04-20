from __future__ import annotations

import xml.etree.ElementTree as ET

from .tlic_models import BranchOptions, LineSection, Point, ProjectData, Structure


def save_project_xml(path: str, project: ProjectData) -> None:
    # Save exactly the state needed to reopen a working project:
    # - branch options (bus/kV/base/temp/etc.)
    # - saved line sections
    # - custom structures created/edited by the user
    #
    # We serialize dataclass fields explicitly so future field additions can be
    # handled in one place and unknown fields can be ignored on load.
    root = ET.Element("TLICProject")

    opts = ET.SubElement(root, "BranchOptions")
    for key, value in project.options.__dict__.items():
        opts.set(key, str(value))

    sections = ET.SubElement(root, "LineSections")
    for s in project.sections:
        ET.SubElement(
            sections,
            "LineSection",
            {
                "cond_name": s.cond_name,
                "static_name": s.static_name,
                "struct_name": s.struct_name,
                "mileage": str(s.mileage),
                "is_custom_structure": str(s.is_custom_structure),
                "mot": str(s.mot),
            },
        )

    customs = ET.SubElement(root, "CustomStructures")
    for name, st in project.custom_structures.items():
        node = ET.SubElement(customs, "Structure", {"name": name})
        for idx, p in enumerate(st.a):
            ET.SubElement(node, "A", {"idx": str(idx), "x": str(p.x), "y": str(p.y)})
        for idx, p in enumerate(st.g):
            ET.SubElement(node, "G", {"idx": str(idx), "x": str(p.x), "y": str(p.y)})

    tree = ET.ElementTree(root)
    tree.write(path, encoding="utf-8", xml_declaration=True)


def load_project_xml(path: str) -> ProjectData:
    # Mirror save_project_xml() schema with defensive parsing:
    # - tolerate missing nodes/attributes
    # - cast by existing BranchOptions field type
    # - skip unknown attributes for forward/backward compatibility
    tree = ET.parse(path)
    root = tree.getroot()

    opts_node = root.find("BranchOptions")
    options = BranchOptions()
    if opts_node is not None:
        for k, v in opts_node.attrib.items():
            if not hasattr(options, k):
                continue
            current = getattr(options, k)
            # Cast from XML string back into the target field type.
            if isinstance(current, bool):
                setattr(options, k, v.lower() in {"1", "true", "yes"})
            elif isinstance(current, int):
                setattr(options, k, int(float(v)))
            elif isinstance(current, float):
                setattr(options, k, float(v))
            else:
                setattr(options, k, v)

    sections: list[LineSection] = []
    for s in root.findall("./LineSections/LineSection"):
        sections.append(
            LineSection(
                cond_name=s.attrib.get("cond_name", ""),
                static_name=s.attrib.get("static_name", ""),
                struct_name=s.attrib.get("struct_name", ""),
                mileage=float(s.attrib.get("mileage", "0") or 0.0),
                is_custom_structure=s.attrib.get("is_custom_structure", "False").lower() in {"1", "true", "yes"},
                mot=float(s.attrib.get("mot", "125") or 125.0),
            )
        )

    customs: dict[str, Structure] = {}
    for node in root.findall("./CustomStructures/Structure"):
        name = node.attrib.get("name", "Custom")
        # Keep fixed point counts so structure consumers can index safely:
        # A -> 3 phase points, G -> up to 2 static points.
        a = [Point(), Point(), Point()]
        g = [Point(), Point()]
        for p in node.findall("A"):
            idx = int(p.attrib.get("idx", "0"))
            if 0 <= idx < 3:
                a[idx] = Point(float(p.attrib.get("x", "0")), float(p.attrib.get("y", "0")))
        for p in node.findall("G"):
            idx = int(p.attrib.get("idx", "0"))
            if 0 <= idx < 2:
                g[idx] = Point(float(p.attrib.get("x", "0")), float(p.attrib.get("y", "0")))
        customs[name] = Structure(name=name, a=a, g=g)

    return ProjectData(options=options, sections=sections, custom_structures=customs)
