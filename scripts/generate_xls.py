# @copyright Copyright (c) 2021 Nicolas GRELLETY
# @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
# @link https://gitlab.internal.lan/octo-project/octowriter
# @link https://github.com/nillyr/octowriter
# @since 0.1.0

from itertools import chain
from pathlib import Path
import re
import shutil
from typing import List
import zipfile

import configparser
import xlsxwriter

from octoconf.entities.baseline import Baseline
from octoconf.entities.category import Category
from octoconf.entities.rule import Rule
import octoconf.utils.global_values as global_values
from octoconf.utils.timestamp import today
import octoconf.utils.config as config
from octoconf.__init__ import __version__, __url__


class XLSGenerator:
    wb: xlsxwriter.workbook.Workbook = None
    _formats: dict = {}

    def __init__(self) -> None:
        pass

    def _add_new_format(self, name: str, values: dict) -> None:
        # fmt:off
        self._formats[name] = self.wb.add_format(
            {
                "bold": values["bold"],
                "border": values["border"],
                "align": values["align"],
                "valign": values["valign"],
                "font_color": values["font_color"],
                "bg_color": values["bg_color"]
            })

    def _get_format(self, name: str) -> xlsxwriter.workbook.Format:
        if name in self._formats:
            return self._formats.get(name)

    def _init_all_format(self):
        self._add_new_format(
            "information_header",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_size": 14,
                "font_color": config.get_config("report_colors", "header_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "header_background_color"
                ),
            },
        )
        self._add_new_format(
            "header",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("report_colors", "header_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "header_background_color"
                ),
            },
        )
        self._add_new_format(
            "sub_header",
            {
                "bold": 0,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("report_colors", "header_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "sub_header_background_color"
                ),
            },
        )
        self._add_new_format(
            "minimal",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("level_colors", "lvl_minimal"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "intermediary",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("level_colors", "lvl_intermediary"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "enhanced",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("level_colors", "lvl_enhanced"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "high",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("level_colors", "lvl_high"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "check",
            {
                "bold": 0,
                "border": 1,
                "align": "left",
                "valign": "vcenter",
                "font_color": config.get_config("report_colors", "default_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "success",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("status_colors", "success"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "failed",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("status_colors", "failed"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "na",
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config("status_colors", "to_be_defined"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "bold",
            {
                "bold": 1,
                "border": 1,
                "align": "left",
                "valign": "vcenter",
                "font_color": config.get_config("report_colors", "default_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "regular",
            {
                "bold": 0,
                "border": 1,
                "align": "left",
                "valign": "vcenter",
                "font_color": config.get_config("report_colors", "default_font_color"),
                "bg_color": config.get_config(
                    "report_colors", "default_background_color"
                ),
            },
        )
        self._add_new_format(
            "classification",
            {
                "bold": 0,
                "border": 1,
                "align": "left",
                "valign": "vcenter",
                "font_color": config.get_config(
                    "classification", "classification_font_color"
                ),
                "bg_color": config.get_config(
                    "classification", "classification_background_color"
                ),
            },
        )
        self._add_new_format(
            "classification_center",
            {
                "bold": 0,
                "border": 0,
                "align": "center",
                "valign": "vcenter",
                "font_color": config.get_config(
                    "classification", "classification_font_color"
                ),
                "bg_color": config.get_config(
                    "classification", "classification_background_color"
                ),
            },
        )

    def _add_conditional_formatting(
        self, ws: xlsxwriter.workbook.Worksheet, range
    ) -> None:
        # fmt:off
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("minimal"),
            'format': self._get_format("minimal")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("intermediary"),
            'format': self._get_format("intermediary")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("enhanced"),
            'format': self._get_format("enhanced")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("high"),
            'format': self._get_format("high")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("success"),
            'format': self._get_format("success")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("failed"),
            'format': self._get_format("failed")
        })
        ws.conditional_format(range, {
            'type': 'text',
            'criteria': 'containing',
            'value': global_values.localize.gettext("na"),
            'format': self._get_format("na")
        })
        # fmt:on

    def _write_results_on_worksheet(
        self, ws: xlsxwriter.workbook.Worksheet, rules: List[Rule]
    ) -> None:
        checkpoint_row = 4
        ws.write(
            f"B{checkpoint_row}",
            global_values.localize.gettext("level"),
            self._get_format("sub_header"),
        )
        ws.merge_range(
            f"C{checkpoint_row}:E{checkpoint_row}",
            "Rule",
            self._get_format("sub_header"),
        )
        ws.write(
            f"F{checkpoint_row}",
            global_values.localize.gettext("result"),
            self._get_format("sub_header"),
        )

        check_row = checkpoint_row + 1
        for rule in rules:
            ws.data_validation(
                f"B{check_row}",
                {
                    "validate": "list",
                    "source": [
                        global_values.localize.gettext("minimal"),
                        global_values.localize.gettext("intermediary"),
                        global_values.localize.gettext("enhanced"),
                        global_values.localize.gettext("high"),
                    ],
                },
            )
            ws.data_validation(
                f"F{check_row}",
                {
                    "validate": "list",
                    "source": [
                        global_values.localize.gettext("success"),
                        global_values.localize.gettext("failed"),
                        global_values.localize.gettext("na"),
                    ],
                },
            )

            ws.write(
                f"B{check_row}",
                global_values.localize.gettext(rule.level),
                self._get_format(rule.level),
            )
            ws.merge_range(
                f"C{check_row}:E{check_row}", rule.title, self._get_format("check")
            )

            key = "success" if rule.compliant == True else "failed"
            ws.write(
                    f"F{check_row}",
                    global_values.localize.gettext(key),
                    self._get_format(key),
            )

            check_row += 1
            checkpoint_row = check_row

    def _write_results(self, categories: List[Category]) -> None:
        for category in categories:
            # It is not possible to use a worksheet's title > 31 chars, so we need to slice
            regex = r"(</?x>)|[^a-zàâçéèêëîïôûù0-9\s\-]"
            category_name = re.sub(
                regex,
                "",
                category.name[0:31],
                0,
                re.IGNORECASE,
            )
            ws = self.wb.add_worksheet(name=category_name)
            ws.hide_gridlines(2)
            ws.set_column("A:A", 2)
            ws.set_column("B:B", 20)
            ws.set_column("C:E", 35)
            ws.set_column("F:F", 20)

            ws.merge_range("C1:E1", "", self._get_format("classification_center"))
            ws.write_formula("C1:E1", "=%s!D10" % (global_values.localize.gettext("information")), self._get_format("classification_center"), "")

            ws.set_row(2, 25)
            ws.merge_range("B3:F3", category.name, self._get_format("header"))
            # Column 'A' (level)
            range_a = xlsxwriter.utility.xl_range(2, 1, 1048575, 1)
            self._add_conditional_formatting(ws, range_a)
            # Column 'E' (result)
            range_e = xlsxwriter.utility.xl_range(0, 5, 1048575, 5)
            self._add_conditional_formatting(ws, range_e)
            # Write results in the worksheet and get nb of success/failed for stacked chart
            self._write_results_on_worksheet(ws, category.rules)

    def _add_charts(self,
        ws: xlsxwriter.worksheet.Worksheet,
        last_row: int) -> None:
    #fmt:on
        """
        A picture is worth a thousand words, and this method generates charts indicating the coverage level of security configurations.
        """
        ws_name = ws.get_name()
        staked_chart_by_lvl = self.wb.add_chart({'type': 'column', 'subtype': 'stacked'})
        staked_chart_by_lvl.set_title({'name': global_values.localize.gettext("compliance_chart_title")})
        staked_chart_by_lvl.set_x_axis({'name': global_values.localize.gettext("levels")})
        staked_chart_by_lvl.set_y_axis({'name': global_values.localize.gettext("nb_checks"), 'major_gridlines': {'visible': False}})

        staked_chart_by_lvl.add_series({
            "name":         f"={ws_name}!$E$4",
            "categories":   f"={ws_name}!$E$5:$H$5",
            "values":       f"={ws_name}!$E${last_row}:H${last_row}",
            "data_labels":  {"value": True},
            "fill":         {"color": "#"+config.get_config("status_colors", "success")},
            "gap":          20
        })

        staked_chart_by_lvl.add_series({
            "name":         f"={ws_name}!$I$4",
            "categories":   f"={ws_name}!$L$5:$L$5",
            "values":       f"={ws_name}!$I${last_row}:L${last_row}",
            "data_labels":  {"value": True},
            "fill":         {"color": "#"+config.get_config("status_colors", "failed")},
            "gap":          20
        })

        # Do not stick the chart on the far left
        ws.insert_chart(f"E{last_row+5}", staked_chart_by_lvl)

    def _add_synthesis_worksheet(self, categories: List[Category]) -> None:
        """
        Resumes all the sheets (categories) of the excel file in order to present in the same sheet the synthesis of the results.
        """
        ws = self.wb.add_worksheet(name=global_values.localize.gettext("summary"))

        ws.hide_gridlines(2)
        ws.set_column("A:A", 2)
        ws.set_column("B:L", 20)
        ws.set_row(2, 25)

        ws.merge_range("B1:L1", "", self._get_format("classification_center"))
        ws.write_formula("B1:L1", "=%s!D10" % (global_values.localize.gettext("information")), self._get_format("classification_center"), "")

        ws.merge_range("B3:L3", global_values.localize.gettext("summary"), self._get_format("header"))
        ws.merge_range(
            "B4:D5", global_values.localize.gettext("categories"), self._get_format("sub_header")
        )
        ws.merge_range(
            "E4:H4", global_values.localize.gettext("success"), self._get_format("sub_header")
        )
        ws.merge_range(
            "I4:L4", global_values.localize.gettext("failed"), self._get_format("sub_header")
        )
        ws.write("E5", global_values.localize.gettext("minimal"), self._get_format("sub_header"))
        ws.write("F5", global_values.localize.gettext("intermediary"), self._get_format("sub_header"))
        ws.write("G5", global_values.localize.gettext("enhanced"), self._get_format("sub_header"))
        ws.write("H5", global_values.localize.gettext("high"), self._get_format("sub_header"))
        ws.write("I5", global_values.localize.gettext("minimal"), self._get_format("sub_header"))
        ws.write("J5", global_values.localize.gettext("intermediary"), self._get_format("sub_header"))
        ws.write("K5", global_values.localize.gettext("enhanced"), self._get_format("sub_header"))
        ws.write("L5", global_values.localize.gettext("high"), self._get_format("sub_header"))

        row = 5
        for category in categories:
            row += 1
            # A = 0, B = 1, C =2, D = 3
            # E = 4, F = 5, G = 6, H = 7
            # I = 8, J = 9, K = 10
            ws.merge_range(
                xlsxwriter.utility.xl_range(row - 1, 1, row - 1, 3),
                category.name,
                self._get_format("check"),
            )
            regex = r"(</?x>)|[^a-zàâçéèêëîïôûù0-9\s\-]"
            category_name = re.sub(
                regex,
                "",
                category.name[0:31],
                0,
                re.IGNORECASE,
            )
            lvl_range = f"'{category_name}'!{xlsxwriter.utility.xl_range(0, 1, 1048575, 1)}"
            results_range = f"'{category_name}'!{xlsxwriter.utility.xl_range(0, 5, 1048575, 5)}"

            levels = [
                f"{lvl_range};\"={global_values.localize.gettext('minimal')}\"",
                f"{lvl_range};\"={global_values.localize.gettext('intermediary')}\"",
                f"{lvl_range};\"={global_values.localize.gettext('enhanced')}\"",
                f"{lvl_range};\"={global_values.localize.gettext('high')}\""
            ]

            success = {f"{results_range};\"={global_values.localize.gettext('success')}\"": levels}
            failed = {f"{results_range};\"={global_values.localize.gettext('failed')}\"": levels}

            start, stop = (4, 8)
            for criteria in success:
                for col in range(start, stop):
                    # Adding "" as the last parameter force the recalculation on file open
                    # see: https://xlsxwriter.readthedocs.io/faq.html
                    ws.write_formula(
                        xlsxwriter.utility.xl_rowcol_to_cell(row - 1, col),
                        f"=COUNTIFS({success[criteria][col - start]}; {criteria})",
                        self._get_format("check"),
                        ""
                    )
            start, stop = (stop, 12)
            for criteria in failed:
                for col in range(start, stop):
                    # Adding "" as the last parameter force the recalculation on file open
                    # see: https://xlsxwriter.readthedocs.io/faq.html
                    ws.write_formula(
                        xlsxwriter.utility.xl_rowcol_to_cell(row - 1, col),
                        f"=COUNTIFS({failed[criteria][col - start]}; {criteria})",
                        self._get_format("check"),
                        ""
                    )

        # Get total values
        ws.merge_range(
                xlsxwriter.utility.xl_range(row, 1, row, 3),
                "Total",
                self._get_format("bold"),
        )
        start_row = 4
        for col in range(start_row, stop):
            ws.write_formula(
                xlsxwriter.utility.xl_rowcol_to_cell(row, col),
                "=SUM(%s)" % (xlsxwriter.utility.xl_range(start_row, col, row - 1, col)),
                self._get_format("bold"),
            )

        self._add_charts(ws, row + 1)

    def _add_information_worksheet(self, baseline_title: str, report_information: dict) -> None:
        ws = self.wb.add_worksheet(name=global_values.localize.gettext("information"))
        ws.hide_gridlines(2)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 100)

        ws.data_validation(f"D10", {
            'validate': 'list',
            'source': [option.lstrip()[0:31] for option in config.get_config("classification", "classification_options").split(",")]
        })

        # Title
        ws.merge_range("B2:D7", global_values.localize.gettext("information_header_title"), self._get_format("information_header"))

        # Title
        ws.merge_range("B9:D9", global_values.localize.gettext("data_classification"), self._get_format("sub_header"))
        # Key
        ws.merge_range("B10:C10", global_values.localize.gettext("classification_level"), self._get_format("bold"))
        # Value
        if "classification-level" in report_information:
            ws.write("D10", report_information["classification-level"], self._get_format("classification"))
        else:
            ws.write("D10", "FIXME", self._get_format("classification"))

        # Title
        ws.merge_range("B12:D12", global_values.localize.gettext("general_information"), self._get_format("sub_header"))
        # Key
        ws.merge_range("B13:C13", global_values.localize.gettext("date_of_completion"), self._get_format("bold"))
        # Value
        ws.write("D13", today(), self._get_format("regular"))
        # Key
        ws.merge_range("B14:C14", global_values.localize.gettext("used_baseline"), self._get_format("bold"))
        # Value
        ws.write("D14", baseline_title, self._get_format("regular"))
        # Key
        ws.merge_range("B15:C15", global_values.localize.gettext("tool_version"), self._get_format("bold"))
        # Value
        ws.write("D15", __version__, self._get_format("regular"))
        # Key
        ws.merge_range("B16:C16", global_values.localize.gettext("online_tool_version"), self._get_format("bold"))
        # Value
        ws.write("D16", __url__, self._get_format("regular"))

        # Title
        ws.merge_range("B18:D18", global_values.localize.gettext("audited_asset"), self._get_format("sub_header"))
        # Key
        ws.merge_range("B19:C19", global_values.localize.gettext("asset"), self._get_format("bold"))
        # Value
        if "audited-asset" in report_information:
            ws.write("D19", report_information["audited-asset"], self._get_format("regular"))
        else:
            ws.write("D19", "FIXME", self._get_format("regular"))

    def _extract_files_from_xlsx(self, xlsx_file, xlsx_folder) -> None:
        with zipfile.ZipFile(xlsx_file, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            for file_name in file_list:
                zip_ref.extract(file_name, path=xlsx_folder)

    def _create_xlsx_from_folder(self, xlsx_folder, output_file) -> None:
        with zipfile.ZipFile(Path(output_file), 'w') as zip_ref:
            xlsx_folder_path = Path(xlsx_folder)
            xml_or_rels_files = chain(xlsx_folder_path.rglob('*.xml'), xlsx_folder_path.rglob('*.rels'))
            for xml_or_rels_file_path in xml_or_rels_files:
                rel_path = xml_or_rels_file_path.relative_to(xlsx_folder_path)
                zip_ref.write(xml_or_rels_file_path, arcname=rel_path)

    def _remove_folder(self, folder_path: Path) -> None:
        folder_path = Path(folder_path)

        if folder_path.is_file():
            folder_path.unlink()
        elif folder_path.is_dir():
            shutil.rmtree(folder_path)
        else:
            print(f"[x] Error: '{folder_path}' is neither a file nor a directory.")

    def _replace_chars(self, input_str: str) -> str:
        input_str = input_str.replace(';', ',')
        input_str = input_str.replace("'", '&apos;')
        input_str = input_str.replace('"', '&quot;')
        input_str = input_str.replace(', &apos;', ',&apos;')
        input_str = input_str.replace(', &quot;', ',&quot;')
        return input_str

    def _format_formulae_for_ms_excel(self, xlsx_folder: Path) -> None:
        regex = r"<f>[a-z0-9\.]+\(\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*('|\")![A-Z]+[0-9]*:[A-Z]+[0-9]*;\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*\s?('|\");\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*('|\")![A-Z]+[0-9]*:[A-Z]+[0-9]*;\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*\s?('|\")\)</f><v></v>"

        # sheet2 = synthesis sheet with all the formulae
        with open(f"{xlsx_folder}/xl/worksheets/sheet2.xml", "r") as sheet2:
            og_sheet2_content = sheet2.read()

        start_index = 0
        formatted_sheet2_content = ""
        matches = re.finditer(regex, og_sheet2_content, re.MULTILINE | re.IGNORECASE)
        for _, match in enumerate(matches, start=1):
            formatted_sheet2_content += og_sheet2_content[start_index:match.start()]
            formatted_sheet2_content += self._replace_chars(match.group())
            start_index = match.end()

        formatted_sheet2_content += og_sheet2_content[start_index:]
        formatted_sheet2_path = Path(f"{xlsx_folder}/xl/worksheets/sheet2.xml")
        formatted_sheet2_path.write_text(formatted_sheet2_content)

    def _generate_microsoft_excel_file(self, input_file: Path) -> None:
        extract_dir = input_file.parent / f"{input_file.stem}_extract"
        output_file = input_file.parent / f"{input_file.stem}-ms-excel-compatible.xlsx"
        self._extract_files_from_xlsx(input_file, extract_dir)
        self._format_formulae_for_ms_excel(extract_dir)
        self._create_xlsx_from_folder(extract_dir, output_file)
        self._remove_folder(extract_dir)

    def generate_xls(self, filename: str, results: Baseline, output_dir: Path, ini_file: Path = None) -> None:
        report_information: dict = dict()
        if ini_file:
            cfg_parser = configparser.ConfigParser()
            cfg_parser.read(ini_file)

            report_information["audited-asset"] = cfg_parser.get("DEFAULT", "audited_asset")
            report_information["classification-level"] = cfg_parser.get("DEFAULT", "classification_level")

        self.wb = xlsxwriter.Workbook(f"{output_dir / filename}.xlsx")
        self._init_all_format()

        self._add_information_worksheet(results.title, report_information)
        self._add_synthesis_worksheet(results.categories)
        self._write_results(results.categories)

        self.wb.close()
        self.wb = None

        self._generate_microsoft_excel_file(Path(f"{output_dir / filename}.xlsx"))
