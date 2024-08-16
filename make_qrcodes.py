#!/bin/env python3
import copy
import functools
import hashlib
import json
import logging
import os.path
import pathlib
import re
import subprocess

import PyPDF2
import click
import semantic_version
from box import Box
from ezodf.sheets import Sheets
from gspread.utils import ExportFormat
from relatorio.templates.opendocument import Template
import yaml
from pandas_ods_reader import read_ods
import ezodf
from logging import getLogger, basicConfig
import yaml
import upload_build
import gspread

basicConfig(level=logging.DEBUG)
logger = getLogger(__name__)

DEFAULT_VERSION = "1.0.0"

# ADD in the config if the pdf must be formatted for pagination (fill up to multiple of 4 pages

class Maker:
    def __init__(self, config_path, target=None, debug=False, upload=False):
        self.config = config = Box.from_yaml(filename=config_path, Loader=yaml.FullLoader)

        # self.template_contents = {}
        # for key, path in self.template_paths.items():
        #     with open(path, "r") as f:
        #         self.template_contents[key] = f.read()

    def _download_from_gdrive(self):
        source_url = self.config.input.gdrive.url
        target_file = self.config.input.local.path
        assert source_url
        assert target_file
        logger.info(f"Download spreadsheet from '{source_url}'")

        gc = gspread.oauth(
            credentials_filename=self.config.input.gdrive.credentials,
            authorized_user_filename=self.config.input.grive.access_token
        )

        gdoc = gc.open_by_url(source_url)
        gcontent = gdoc.export(ExportFormat.OPEN_OFFICE_SHEET)

        with open(target_file, "wb") as f:
            f.write(gcontent)

    def make(self):

        if self.source_url and (self.always_download or not os.path.exists(self.source_path)):
            self._download_from_gdrive(self.source_url, self.source_path)

        characters_context = self.load_ods_to_dict(self.source_path)

        if not characters_context:
            logger.warning("Empty characters context, nothing to do")
            exit(1)

        logger.info(f"Create output directory at '{self.output_path}'")
        pathlib.Path(self.output_path).mkdir(exist_ok=True)

        logger.info(f"Load document template from '{self.template_paths}'")

        from genshi.template import MarkupTemplate

        for ext in ("pdf", "html", "json", "yaml"):
            pathlib.Path(f"./build/{ext}/").mkdir(exist_ok=True)

        templates = {}

        for key, template_content in self.template_contents.items():
            templates[key] = MarkupTemplate(template_content, encoding="utf-8")

        for filename, values in characters_context.items():
            if self.target_character and filename != self.target_character:
                continue

            logger.info(f"Process character {filename}")
            is_png = filename.startswith("PNG")

            if not values or not values["pg"]["nome"]:
                logger.warning(f"Empty values for {filename}")
                continue

            Box(values).to_json(os.path.join(self.output_path, "json", f"{filename}.json"), indent=2)
            Box(values).to_yaml(os.path.join(self.output_path, "yaml", f"{filename}.yaml"))

            logger.debug(f"Renderize {filename}")
            values["debug"] = self.debug

            version_cache = self.versions.get(filename)
            current_semver = semantic_version.Version(version_cache.get("version", DEFAULT_VERSION))
            version = str(current_semver)
            info = {
                "title": filename,
                "version": version
            }

            for key, template in templates.items():
                generated = (template.generate(o=values, i=info)
                             .render("xhtml", encoding="utf-8"))

                html_path = os.path.join(self.output_path, "html", f"{key}-{filename}.html")
                logger.info(f"Write output to '{html_path}'")

                with open(html_path, "wb") as f:
                    f.write(generated)

                logger.info(f"Convert '{html_path}' to pdf")
                pdf_path = os.path.join(self.output_path, "pdf", f"{key}-{filename}.pdf")
                os.system(
                    f"{CHROME_BROWSER} --headless --run-all-compositor-stages-before-draw --print-to-pdf-no-header --print-to-pdf='{pdf_path}' '{html_path}'")

                # fix number of pages
                if key == "tutto":
                    with open(pdf_path, 'rb') as a:
                        pdf = PyPDF2.PdfFileReader(a)
                        outPdf = PyPDF2.PdfFileWriter()
                        outPdf.appendPagesFromReader(pdf)
                        numPages = pdf.getNumPages()
                        if numPages % 4 != 0:
                            numPagesToAdd = 4 - (numPages % 4)
                            for i in range(numPagesToAdd):
                                outPdf.addBlankPage()

                        with open(pdf_path + ".fix", 'wb') as out:
                            outPdf.write(out)

                        os.system(f"mv -v '{pdf_path}.fix' '{pdf_path}'")

            # upload
        if self.upload:
            upload_build.upload()
        else:
            logger.info("Skip upload")

    def load_ods_to_dict(self, source_odt):
        logger.info(f"Load source spreadsheet from '{source_odt}'")
        spreadsheet = ezodf.opendoc(source_odt)
        ezodf.config.reset_table_expand_strategy()
        sheets: Sheets = spreadsheet.sheets

        context = Box()

        logger.info(f"Load common references")
        sheet = sheets['.config_refs']
        self.load_common_field_references(sheet)
        sheet = sheets['.config_replace']
        self.load_important_words(sheet)

        logger.info(f"Load common default data")
        with open(self.default_character_path) as f:
            self.default_character_data = yaml.full_load(f)
        sheet = sheets['.scheda_comune']
        self.default_character_data['root_path'] = "char_maker"
        self.load_sheet_to_dict(self.default_character_data, sheet)

        # logger.info(f"Load hidden handouts")
        # sheet = sheets['.handouts']
        #
        # tmp = copy.deepcopy(self.default_character_data)
        # self.load_sheet_to_dict(tmp, sheet)  # load references only

        for sheet in [s for s in sheets if not s.name.startswith(".")]:
            if self.target_character is not None and sheet.name != self.target_character:
                continue

            with open(self.default_character_path) as f:
                copy_default = not sheet.name.startswith("PNG")
                if copy_default:
                    context[sheet.name] = copy.deepcopy(self.default_character_data)
                else:
                    context[sheet.name] = {}
            data = context[sheet.name]
            self.load_sheet_to_dict(data, sheet)

        return context

    def load_common_field_references(self, sheet):
        headers = {x.value: i for i, x in enumerate(sheet.row(0)) if x.value}
        for row in list(sheet.rows())[1:]:
            values = dict(zip(headers.keys(), [r.value for r in row]))
            if values['riferimento']:
                self.reference_translations[values['riferimento']] = values['campo']

    def load_important_words(self, sheet):
        headers = {x.value: i for i, x in enumerate(sheet.row(0)) if x.value}
        for row in list(sheet.rows())[1:]:
            values = dict(zip(headers.keys(), [r.value for r in row]))
            patt = values['pattern']
            if not patt:
                break
            if values.get("enable", "TRUE") == "FALSE":
                continue
            repl = values['replace']
            css_class = values['css_class']
            if not repl:
                repl = r"\g<0>"
            if css_class:
                repl = rf"<span class='{css_class}'>{repl}</span>"
            self.text_replacements[patt] = repl

    def load_sheet_to_dict(self, data, sheet):
        logger.info(f"Load sheet {sheet.name}")

        headers = {x.value: i for i, x in enumerate(sheet.row(0)) if x.value}

        data["note_regia"] = []

        version_cache = self.versions.get(sheet.name, {})
        semver = semantic_version.Version(version_cache.get("version", DEFAULT_VERSION))
        current_patch_hash = version_cache.get("hash_patch", None)  # se cambia aumenta di uno il patch
        current_minor_hash = version_cache.get("hash_minor", None)  # se cambia aumenta il minor

        new_patch_hash = hashlib.sha1(
            json.dumps([[x.value for x in row[1:]] for row in sheet.rows()]).encode()).hexdigest()
        new_minor_hash = hashlib.sha1(json.dumps([row[0].value for row in sheet.rows()]).encode()).hexdigest()
        if current_minor_hash is not None and new_minor_hash != current_minor_hash:
            semver.minor = semver.minor + 1
            semver.patch = 0
            logger.info(f"New minor version {semver}")
        elif current_patch_hash is not None and new_patch_hash != current_patch_hash:
            semver.patch = semver.patch + 1
            logger.info(f"New patch version {semver}")

        version = str(semver)
        version_cache['version'] = version
        version_cache['hash_patch'] = new_patch_hash
        version_cache['hash_minor'] = new_minor_hash
        self.versions[sheet.name] = version_cache
        self.write_versions()

        for row in list(sheet.rows())[1:]:
            values = dict(zip(headers.keys(), [r.value for r in row]))
            try:
                self.load_row_to_dict(data, values)
            except Exception as e:
                print(f"Error loading line: {e}")

    def load_row_to_dict(self, data, values):

        if not values:
            return False

        field = values['campo']
        if not field:
            return False

        ref_unique = values.get('ref_univoco', None)

        # if field in ("DEL","del","REMOVE","remove","EXCLUDE","exclude"):
        #     assert ref_unique is not None
        #     logger.debug(f"Remove {ref_unique}")
        #     self.field_operation(data, ref_unique, None, None, "del")
        #     continue

        if field.startswith("!"):
            field = field[1:]
            logger.debug(f"Remove {field}")
            self.field_operation(data, field, None, None, "del")
            return True

        titolo = values.get('titolo')
        if titolo is not None:
            if isinstance(titolo, str) and titolo.startswith("$"):
                target = titolo[1:]
                logger.debug(f"Include {target}")
                if not target in self.reference_data:
                    raise ValueError(f"Invalid reference {target}")
                value = self.reference_data[target]
                self.field_operation(data, field, value, ref_unique, "set")
                return True

        if field.startswith("$"):
            target = field[1:]
            logger.debug(f"Include {target}")
            if not target in self.reference_data:
                raise ValueError(f"Invalid reference {target}")
            value = self.reference_data[target]
            field = re.sub(r'(.*)\[.*\]', r'\1[]', self.query_ref_uniques(target))
            self.field_operation(data, field, value, ref_unique, "set")
            return True

        # if field in ignore_fields:
        #     logger.debug(f"Ignore field {field}")
        #     return False

        image = None
        if values['inclusione_file']:
            image = os.path.join('assets', 'images', values['inclusione_file'])

        import markdown
        for key, val in values.items():
            if key in ("titolo", "contenuto", "importante", "note_regia"):
                if val and isinstance(val, str):
                    val = markdown.markdown(val)
                    val = re.sub("(^<p>|</p>$)", "", val, flags=re.IGNORECASE)
                    val = re.sub("</p>\s*<p>", "<br/>", val, flags=re.IGNORECASE)
                    if "???" in val:
                        val = f"<div style='color:red'>{val}</div>"
                    for patt, repl in self.text_replacements.items():
                        val = re.sub(patt, repl, val)
                    val = val.encode('ascii', 'xmlcharrefreplace').decode('ascii')
                    values[key] = val

        field_content = Box({
            'titolo': values.get('titolo', ""),
            'contenuto': values.get('contenuto', ""),
            'importante': values.get('importante', ""),
            'file': None,
            'css': values.get('css', ""),
            'css_class': values.get('css_class', ""),
            'note_regia': values.get('note_regia', ""),
            'field': field,
            'unique': ref_unique
        })

        if image:
            field_content['file'] = image

        if field_content['note_regia']:
            data["note_regia"].append(field_content)

        # se un solo elemento disponibile, toglie il dizionario
        content_data = (
            field_content['titolo'], field_content['contenuto'], field_content['importante'], field_content['file']
        )
        one_field = functools.reduce(lambda a, b: -1 if (a and b) else (a if a else b), content_data)

        if isinstance(one_field, str):
            field_content['contenuto'] = one_field
        if all(not x for x in content_data):
            field_content = None

        if ref_unique:
            self.reference_data[ref_unique] = field_content

        self.field_operation(data, field, field_content, ref_unique, "set")
        return True

    def query_ref_uniques(self, current_selector):
        if current_selector in self.reference_translations:
            old_selector = current_selector
            current_selector = self.reference_translations[old_selector]
            logger.debug(f"selector {old_selector} -> {current_selector}")
        return current_selector

    def field_operation(self, parent, main_selector, value, ref_unique=None, command=None,
                        hidden=False):

        main_selector = self.query_ref_uniques(main_selector)

        tokens = re.split(r'\.', main_selector, 1)
        final_selector = []

        while tokens:
            try:
                current_selector = tokens.pop(0)
                current_selector = self.query_ref_uniques(current_selector)

                token_is_list = current_selector.endswith("]")
                token_is_last = not tokens

                current_selector, *other = re.split(r'\[|]', current_selector)
                human_selector_num = None
                current_selector_num = None
                current_selector = current_selector
                current_selector = self.query_ref_uniques(current_selector)

                if other and other[0]:
                    human_selector_num = int(other[0])
                    current_selector_num = human_selector_num if (human_selector_num < 0) else (human_selector_num - 1)

                if isinstance(parent, (Box, dict)):
                    if current_selector not in parent:
                        if token_is_list:
                            parent[current_selector] = []
                        elif token_is_last:
                            if command == "get":
                                return parent[current_selector]
                            elif command == "del":
                                parent[current_selector] = {}
                            elif command == "set":
                                parent[current_selector] = value  # end
                            final_selector.append(f"{current_selector}")
                            break
                        else:  # dict
                            parent[current_selector] = {}
                    current = parent[current_selector]
                    if token_is_list:
                        assert isinstance(current, list), f"token is list, but current is not a list: {current}"
                        if current_selector_num is not None:
                            if current_selector_num < 0:
                                current_selector_num = len(current) + current_selector_num
                                human_selector_num = current_selector_num + 1
                            if token_is_last:
                                if command == "get":
                                    return current[current_selector_num]
                                elif command == "del":
                                    current[current_selector_num] = {}
                                elif command == "set":
                                    current[current_selector_num] = value  # end
                                final_selector.append(f"{current_selector}[{human_selector_num}]")
                                break
                            else:
                                parent = current[
                                    current_selector_num]  # go one level deep, can be a dict or simple value
                                final_selector.append(f"{current_selector}[{human_selector_num}]")
                                continue
                        else:
                            if token_is_last:
                                current_selector_num = len(current)
                                current.append(value)  # end
                                human_selector_num = current_selector_num + 1
                                final_selector.append(f"{current_selector}[{human_selector_num}]")
                                break
                            else:
                                current_selector_num = len(current)
                                human_selector_num = current_selector_num + 1
                                current.append(
                                    {'index': human_selector_num})  # append a new dictionary as last element?
                                parent = current[current_selector_num]
                                final_selector.append(f"{current_selector}[{human_selector_num}]")
                                continue
                    elif token_is_last:
                        if command == "get":
                            return parent[current_selector]
                        elif command == "del":
                            parent[current_selector] = {}
                        elif command == "set":
                            parent[current_selector] = value  # end
                        final_selector.append(f"{current_selector}")
                        break
                    else:  # dict
                        assert isinstance(current, (Box, dict))
                        parent = current  # go one level deep
                        final_selector.append(f"{current_selector}")
                        continue

                elif isinstance(parent, list):
                    # non dovremmo MAI arrivarci qua!
                    raise NotImplementedError()

                else:
                    # e neanche qua
                    raise NotImplementedError(f"{parent}")
            except Exception as e:
                logger.error(f"Selector '{main_selector}' while processing '{current_selector}': {e}")
                raise

        if ref_unique:
            final_selector = ".".join(final_selector)
            logger.info(f"Set new unique ref {ref_unique} = {final_selector}")
            self.reference_translations[ref_unique] = final_selector


@click.command()
@click.option('-c', '--config', default='./assets/config.yaml', help='make config')
@click.option('-t', '--target', default=None, help='select a character')
@click.option('-d', '--debug', default=False, help='enable debug', is_flag=True)
def main(config, target, debug):
    Maker(config, target, debug).make()


if __name__ == "__main__":
    main()
