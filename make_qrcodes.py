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
import segno
import semantic_version
from box import Box
from ezodf.sheets import Sheets
from gspread.utils import ExportFormat
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
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


def incremental_index():
    NEXT = 0
    while True:
        yield NEXT
        NEXT += 1


INCREMENTAL_INDEX = incremental_index()


class Maker:
    def __init__(self, config_path, debug=False):
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
        logger.info(f"Download spreadsheet from '{source_url}' to '{target_file}'")

        gc = gspread.oauth(
            credentials_filename=self.config.input.gdrive.credentials,
            authorized_user_filename=self.config.input.gdrive.access_token
        )

        gdoc = gc.open_by_url(source_url)
        gcontent = gdoc.export(ExportFormat.OPEN_OFFICE_SHEET)

        with open(target_file, "wb") as f:
            f.write(gcontent)

    def _load_ods_to_dict(self):
        source_odt = self.config.input.local.path
        logger.info(f"Load spreadsheet from '{source_odt}'")
        spreadsheet = ezodf.opendoc(source_odt)
        ezodf.config.reset_table_expand_strategy()
        sheets: Sheets = spreadsheet.sheets

        context = Box()

        for sheet_name in sheets.names():
            # ignore dotted sheets
            if sheet_name.startswith("."):
                continue
            logger.debug(f"Read sheet: {sheet_name}")

            sheet = sheets[sheet_name]

            headers = None
            for row in sheet.rows():
                if headers is None:
                    headers = [x.value for x in row if x]
                else:
                    line = dict(zip(headers, [x.value for i, x in enumerate(row) if i < len(headers)]))
                    if line["content"] is None:
                        break
                    if line["id"] is None:
                        line["id"] = next(INCREMENTAL_INDEX)
                    if line["count"] is None:
                        line["count"] = 1
                    if line["size"] is None:
                        line["size"] = 5
                    context[line["id"]] = line

        return context

    def make(self):
        if self.config.input.gdrive.download:
            self._download_from_gdrive()

        qrcodes = self._load_ods_to_dict()

        for id, line in qrcodes.items():
            logger.debug(f"Make qrcode: {id}")
            image_path = f'./build/{id}.png'
            qrcode = segno.make_qr(line["content"])
            qrcode.save(image_path, scale=line["size"])
            qrcodes[id]["img"] = f"./{id}.png"

        logger.info(f"Create output at '{self.config.output.local.path}'")
        pathlib.Path(self.config.output.local.path).parent.mkdir(exist_ok=True)

        logger.info(f"Load document template from '{self.config.template.html.path}'")

        from genshi.template import MarkupTemplate
        template = MarkupTemplate(open(self.config.template.html.path, encoding="utf-8"))
        generated = (template.generate(values={"qrcodes": sorted(qrcodes.values(), key=lambda x: x["id"])}).render("xhtml", encoding="utf-8"))

        html_path = self.config.output.local.path.replace(".pdf", ".html")
        logger.info(f"Write output to '{html_path}'")
        with open(html_path, "wb") as f:
            f.write(generated)

        logger.info(f"Convert '{html_path}' to pdf")
        pdf_path = self.config.output.local.path
        os.system(
            f"{self.config.output.local.browser} --headless --run-all-compositor-stages-before-draw --print-to-pdf-no-header --print-to-pdf='{pdf_path}' '{html_path}'"
        )

        if self.config.output.gdrive.upload:
            logger.info("Upload to Google Drive")
            self._upload(pdf_path)


    def _upload(self, file):
        gauth = GoogleAuth()
        gauth.LoadCredentialsFile(self.config.input.gdrive.access_token)
        if gauth.credentials is None:
            gauth.LocalWebserverAuth()
        elif gauth.access_token_expired:
            gauth.Refresh()
        else:
            gauth.Authorize()
        gauth.SaveCredentialsFile(self.config.input.gdrive.access_token)

        drive = GoogleDrive(gauth)

        file_list = drive.ListFile({'q': f"'{self.config.output.gdrive.folder}' in parents and trashed=false"}).GetList()
        file_list = {f['title']: f for f in file_list}
        if file in file_list.keys():
            logger.info(f"Upload and replace {file}")
            gfile = drive.CreateFile({'id': file_list[file]['id']})
            del file_list[file]
        else:
            logger.info(f"Upload {file}")
            gfile = drive.CreateFile({'parents': [{'id': self.config.output.gdrive.folder}]})
        gfile.SetContentFile(file)

        gfile.Upload()

        permission = gfile.InsertPermission({
            'type': 'anyone',
            'value': 'anyone',
            'role': 'reader'})

@click.command()
@click.option('-c', '--config', default='./config.yaml', help='make config')
@click.option('-d', '--debug', default=False, help='enable debug', is_flag=True)
def main(config, debug):
    Maker(config, debug).make()


if __name__ == "__main__":
    main()
