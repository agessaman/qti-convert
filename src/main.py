#!/usr/bin/env python3
"""
QTI to other formats converter
"""

import argparse
import json
import random
from pathlib import Path

from logzero import logger
from lxml import etree

import config
import formats
from qti_parser import assessment, assessment_meta, html_util, item, variant

__author__ = config.__author__
__description__ = config.__description__
__license__ = config.__license__
__version__ = config.__version__


def _form_rng(seed, form_index):
    if seed is not None:
        return random.Random(seed + form_index)
    return random.Random()


def _strip_forms(forms, strip_html):
    if not strip_html:
        return forms
    for form in forms:
        for assessment_data in form["assessment"]:
            html_util.strip_metadata(assessment_data["metadata"])
            assessment_data["question"] = [
                html_util.strip_question(question)
                for question in assessment_data["question"]
            ]
    return forms


def _build_assessments(input_path, xml_resource, args):
    """Parse manifest resource into one or more form payloads."""
    metadata_path = input_path.parent / xml_resource.get("identifier") / "assessment_meta.xml"
    assessment_id = xml_resource.get("identifier")
    assessment_xml = input_path.parent / assessment_id / (assessment_id + ".xml")
    metadata = assessment_meta.get_metadata(metadata_path)

    logger.info(f"this assessment: {assessment_xml}")

    if args.forms == 1:
        blocks = assessment.parse_structure(assessment_xml)
        questions = variant.generate_questions(
            blocks,
            "group",
            random.Random(),
            apply_pools=False,
        )
        return [{
            "label": None,
            "assessment": [{
                "id": assessment_id,
                "metadata": metadata,
                "question": questions,
            }],
        }]

    blocks = assessment.parse_structure(assessment_xml)
    forms = []
    for form_index in range(args.forms):
        label = variant.form_label(form_index)
        rng = _form_rng(args.seed, form_index)
        questions = variant.generate_questions(
            blocks,
            args.shuffle_scope,
            rng,
            apply_pools=not args.output_all,
        )
        form_metadata = dict(metadata)
        form_metadata["form_label"] = label
        forms.append({
            "label": label,
            "assessment": [{
                "id": assessment_id,
                "metadata": form_metadata,
                "question": questions,
            }],
        })
    return forms


def _form_output_path(base_path, label, form_output):
    path = Path(base_path)
    if form_output == "combined":
        return str(path)
    stem = path.stem
    suffix = path.suffix
    safe_label = label.replace(" ", "-")
    return str(path.with_name(f"{stem}-{safe_label}{suffix}"))


def _write_json(forms, args):
    if args.forms == 1:
        payload = {"assessment": forms[0]["assessment"]}
        if args.output:
            logger.info("Writing JSON to '%s'...", args.output)
            with open(args.output, "w") as outfile:
                json.dump(payload, outfile, indent=2)
        else:
            logger.info("Writing JSON to STDOUT...")
            print(json.dumps(payload, indent=2))
        return

    if args.form_output == "combined":
        payload = {"forms": forms}
        outfile = args.output or "output.json"
        logger.info("Writing JSON to '%s'...", outfile)
        with open(outfile, "w") as out:
            json.dump(payload, out, indent=2)
        return

    for form in forms:
        payload = {"forms": [form]}
        if args.output:
            outfile = _form_output_path(args.output, form["label"], "separate")
        else:
            slug = form["label"].lower().replace(" ", "-")
            outfile = f"{slug}.json"
        logger.info("Writing JSON to '%s'...", outfile)
        with open(outfile, "w") as out:
            json.dump(payload, out, indent=2)


def _write_docx(forms, args):
    if args.forms == 1:
        outfile = args.output or "output.docx"
        logger.info("Writing DOCX to '%s'...", outfile)
        formats.docx.write_file(
            {"assessment": forms[0]["assessment"]},
            outfile,
            split_matching=args.split_matching,
            output_key=args.output_key,
            rng=_form_rng(args.seed, 0),
        )
        return

    if args.form_output == "combined":
        outfile = args.output or "output.docx"
        logger.info("Writing DOCX to '%s'...", outfile)
        formats.docx.write_forms(
            forms,
            outfile,
            combined=True,
            split_matching=args.split_matching,
            output_key=args.output_key,
            seed=args.seed,
        )
        return

    for form_index, form in enumerate(forms):
        if args.output:
            outfile = _form_output_path(args.output, form["label"], "separate")
        else:
            safe_label = form["label"].replace(" ", "-")
            outfile = f"output-{safe_label}.docx"
        logger.info("Writing DOCX to '%s'...", outfile)
        formats.docx.write_file(
            {"assessment": form["assessment"]},
            outfile,
            split_matching=args.split_matching,
            output_key=args.output_key,
            rng=_form_rng(args.seed, form_index),
        )


def main(args):
    logger.info(__description__)

    try:
        xml_doc = etree.parse(args.input)
        input_path = Path(args.input)

        all_forms = []
        for xml_resource in xml_doc.getroot().findall(
            ".//{http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1}resource[@type='imsqti_xmlv1p2']"
        ):
            all_forms.extend(_build_assessments(input_path, xml_resource, args))

        all_forms = _strip_forms(all_forms, args.strip_html)

        if args.format.lower() == "json":
            _write_json(all_forms, args)
        elif args.format.lower() == "docx":
            _write_docx(all_forms, args)
        elif args.format.lower() == "pdf":
            logger.error("Format not supported yet: " + args.format)
        else:
            logger.error("Unknown format: " + args.format)

    except OSError as e:
        logger.error("%s", e)
    except etree.ParseError as e:
        logger.error("XML parser error: %s", e)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert QTI files into other formats.", add_help=False)
    parser.add_argument("input", help="QTI input file (imsmanifest.xml).")
    parser.add_argument("-v", action="count", default=0, help="Verbosity (-v, -vv, etc).")
    parser.add_argument("-f", action="store", dest="format", default="json", help="Output format, defaults to JSON.")
    parser.add_argument("-o", action="store", dest="output", help="Output file.")
    parser.add_argument("--forms", type=int, default=1, metavar="N", help="Number of randomized forms to generate (default: 1).")
    parser.add_argument(
        "--shuffle-scope",
        choices=["group", "test"],
        default="group",
        help="Shuffle within question groups (default) or across the entire test.",
    )
    parser.add_argument("--seed", type=int, default=None, help="RNG seed for reproducible form generation.")
    parser.add_argument(
        "--form-output",
        choices=["separate", "combined"],
        default="separate",
        help="Write one file per form (default) or a single combined file.",
    )
    parser.add_argument(
        "--strip-html",
        action="store_true",
        help="Strip HTML tags from question and answer text.",
    )
    parser.add_argument(
        "--split-matching",
        action="store_true",
        help="Split matching questions into even batches of at most 5 for scantron (A-E).",
    )
    parser.add_argument(
        "--output-key",
        action="store_true",
        help="Append an answer key on a separate page (docx only).",
    )
    parser.add_argument(
        "--output-all",
        action="store_true",
        help="Include all questions in each pool when generating multiple forms (ignore selection_number).",
    )
    parser.add_argument("--version", action="version", help="Display version and exit.", version="%(prog)s (version {version})".format(version=__version__))
    parser.add_argument("-h", "--help", action="help", default=argparse.SUPPRESS, help="Show this help message and exit.")
    args = parser.parse_args()
    main(args)
