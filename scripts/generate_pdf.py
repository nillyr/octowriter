# @copyright Copyright (c) 2021-2023 Nicolas GRELLETY
# @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
# @link https://github.com/nillyr/octowriter
# @since 1.0.0b

from pathlib import Path
import platform
import subprocess
from typing import List

from octoconf.__init__ import __version__, __url__

from octoconf.entities.baseline import Baseline
from octoconf.entities.category import Category
from octoconf.entities.rule import Rule

import octoconf.utils.global_values as global_values
from octoconf.utils.timestamp import today


class PDFGenerator:
    def __init__(self) -> None:
        self._template_dir = Path(__file__).resolve().parent.parent / "template"

        self._header_file = self._template_dir / "header.adoc"
        self._introduction_file = self._template_dir / "introduction.adoc"
        self._synthesis_file = self._template_dir / "synthesis.adoc"

    def _is_asciidoctor_pdf_installed(self) -> bool:
        if platform.system() == "Windows":
            cmd = ["powershell.exe", 'if (Get-Command "asciidoctor-pdf") { "true" }']
        else:
            cmd = '[ -x "$(command -v asciidoctor-pdf)" ] && echo "true"'

        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8").strip() == "true"

    def _initialize_report(self, filename: str, baseline_name: str, audited_asset: str) -> dict:
        report_information: dict = dict()

        while True:
            report_information["filename"] = filename

            report_information["document-title"] = global_values.localize.gettext("compliance_report_title")
            report_information["document-subtitle"] = audited_asset

            report_information["auditee_name"] = input(f'{global_values.localize.gettext("auditee_name")} : ')
            report_information["auditee_contact_full_name"] = input(f'{global_values.localize.gettext("auditee_contact_full_name")} : ')
            report_information["auditee_contact_email"] = input(f'{global_values.localize.gettext("auditee_contact_email")} : ')

            report_information["project_manager_full_name"] = input(f'{global_values.localize.gettext("project_manager_full_name")} : ')
            report_information["project_manager_email"] = input(f'{global_values.localize.gettext("project_manager_email")} : ')

            report_information["authors_list_full_name"] = input(f'{global_values.localize.gettext("authors_list_full_name")} : ')
            report_information["authors_list_email"] = input(f'{global_values.localize.gettext("authors_list_email")} : ')

            report_information["baseline_name"] = baseline_name

            report_information["revnumber"] = "1.0"
            report_information["revdate"] = today()

            report_information["confidentiality-level"] = input(f'{global_values.localize.gettext("classification_level")} : ')

            if input(f'{global_values.localize.gettext("init_user_confirmation")} [y/N] : ').upper().strip() == "Y":
                break

        return report_information

    def _include_file_in_header(self, file_to_include: str, build_dir: Path) -> None:
        with open(build_dir / self._header_file.name, "a") as file:
            file.write(f"include::{file_to_include}[]\n")

    def _generate_header_file(self,
                                report_information: dict,
                                build_dir: Path) -> None:
        with open(self._header_file, "r") as file:
            header = file.read()

        header = header.replace("MATCH_AND_REPLACE_DOCUMENT_LANG", global_values.get_locale().upper())
        header = header.replace("MATCH_AND_REPLACE_FILENAME", report_information["filename"])
        header = header.replace("MATCH_AND_REPLACE_DOCUMENT_TITLE", report_information["document-title"])
        header = header.replace("MATCH_AND_REPLACE_DOCUMENT_SUBTITLE", report_information["document-subtitle"])
        header = header.replace("MATCH_AND_REPLACE_AUDITEE_NAME", report_information["auditee_name"])
        header = header.replace("MATCH_AND_REPLACE_AUDITEE_CONTACT_FULL_NAME", report_information["auditee_contact_full_name"])
        header = header.replace("MATCH_AND_REPLACE_AUDITEE_CONTACT_EMAIL", report_information["auditee_contact_email"])
        header = header.replace("MATCH_AND_REPLACE_PROJECT_MANAGER_FULL_NAME", report_information["project_manager_full_name"])
        header = header.replace("MATCH_AND_REPLACE_PROJECT_MANAGER_EMAIL", report_information["project_manager_email"])
        header = header.replace("MATCH_AND_REPLACE_AUTHORS_LIST_FULL_NAME", report_information["authors_list_full_name"])
        header = header.replace("MATCH_AND_REPLACE_AUTHORS_LIST_EMAIL", report_information["authors_list_email"])
        header = header.replace("MATCH_AND_REPLACE_BASELINE_NAME", report_information["baseline_name"])
        header = header.replace("MATCH_AND_REPLACE_REVNUMBER", report_information["revnumber"])
        header = header.replace("MATCH_AND_REPLACE_REVDATE", report_information["revdate"])
        header = header.replace("MATCH_AND_REPLACE_CONFIDENTIALITY_LEVEL", report_information["confidentiality-level"])
        header = header.replace("MATCH_AND_REPLACE_TEMPLATE_DIR", str(self._template_dir))
        header = header.replace("MATCH_AND_REPLACE_REPO_URL", __url__)
        header = header.replace("MATCH_AND_REPLACE_PROJECT_VERSION", __version__)

        with open(build_dir / self._header_file.name, "w") as file:
            file.write(header)

    def _generate_introduction_file(self, authors: dict, auditee: dict, build_dir: Path) -> None:
        with open(self._introduction_file, "r") as file:
            introduction = file.read()

        introduction = introduction.replace("MATCH_AND_REPLACE_PARTICIPANTS", global_values.localize.gettext("participants"))
        introduction = introduction.replace("MATCH_AND_REPLACE_ROLE", global_values.localize.gettext("role"))
        introduction = introduction.replace("MATCH_AND_REPLACE_CONTACT_INFORMATION", global_values.localize.gettext("contact_information"))

        introduction = introduction.replace("MATCH_AND_REPLACE_AUDITEE", global_values.localize.gettext("auditee"))

        auditee_list_str = ""
        for key, value in auditee.items():
            auditee_list_str += f"! *{key}*\n"
            auditee_list_str += f"! {value}\n"
            auditee_list_str += "\n"
        introduction = introduction.replace("MATCH_AND_REPLACE_ARRAY_AUDITEE", auditee_list_str)

        introduction = introduction.replace("MATCH_AND_REPLACE_PROJECT_MANAGEMENT", global_values.localize.gettext("project_management"))
        introduction = introduction.replace("MATCH_AND_REPLACE_AUTHORS", global_values.localize.gettext("authors"))

        authors_list_str = ""
        for key, value in authors.items():
            authors_list_str += f"! *{key}*\n"
            authors_list_str += f"! {value}\n"
            authors_list_str += "\n"
        introduction = introduction.replace("MATCH_AND_REPLACE_ARRAY_AUTHORS", authors_list_str)

        introduction = introduction.replace("MATCH_AND_REPLACE_MODIFICATION_HISTORY", global_values.localize.gettext("modification_history"))
        introduction = introduction.replace("MATCH_AND_REPLACE_AUTHOR", global_values.localize.gettext("author"))
        introduction = introduction.replace("MATCH_AND_REPLACE_REPORT_WRITING", global_values.localize.gettext("report_writing"))

        with open(build_dir / self._introduction_file.name, "w") as file:
            file.write(introduction)

        self._include_file_in_header(str(build_dir / self._introduction_file.name), build_dir)

    def _generate_synthesis_file(self,
                                categories: List[Category],
                                build_dir: Path) -> None:
        with open(self._synthesis_file, "r") as file:
            synthesis = file.read()

        synthesis = synthesis.replace("MATCH_AND_REPLACE_NC_SUMMARY_TITLE", global_values.localize.gettext("nc_summary_title"))

        synthesis = synthesis.replace("MATCH_AND_REPLACE_RULE_NAME", global_values.localize.gettext("rule_name"))
        synthesis = synthesis.replace("MATCH_AND_REPLACE_RULE_LEVEL", global_values.localize.gettext("rule_level"))
        synthesis = synthesis.replace("MATCH_AND_REPLACE_RULE_SEVERITY", global_values.localize.gettext("rule_severity"))

        non_conformity_rows = ""
        for category in categories:
            for rule in category.rules:
                non_conformity_rows += f"| <<{category.category}>> | <<nc_{rule.id}>> | {global_values.localize.gettext(rule.level)} | {global_values.localize.gettext(rule.severity)} \n"

        synthesis = synthesis.replace("MATCH_AND_REPLACE_NON_CONFORMITY", non_conformity_rows)

        with open(build_dir / self._synthesis_file.name, "w") as file:
            file.write(synthesis)

        self._include_file_in_header(str(build_dir / self._synthesis_file.name), build_dir)

    def _generate_rule_file(self, rule: Rule, build_dir: Path) -> None:
        rule_file_content = f"=== {rule.title}\n"
        rule_file_content += f"{rule.description}\n"

        reference_content = ""
        for reference in rule.references:
            reference_content += f"* {reference}\n"

        if reference_content != "":
            rule_file_content += f'\n{global_values.localize.gettext("see_also")}\n\n'
            rule_file_content += reference_content

        rule_file_content += "\n"
        rule_file_content += """.{0}
[source%linenums,shell]
[options="nowrap"]
----
{1}
----\n\n""".format(global_values.localize.gettext("check_command"), rule.check)

        rule_file_content += """.{0}
[source%linenums,console]
[options="nowrap"]
----
{1}
----\n\n""".format(global_values.localize.gettext("expected_result"), rule.expected)

        rule_file_content += """.{0}
[source%linenums,console]
[options="nowrap"]
----
{1}
----\n\n""".format(global_values.localize.gettext("terminal_output"), rule.output)

        if rule.compliant:
            rule_file_content += """ifeval::["{document-lang}" == "EN"]
[.compliant]#The configuration is compliant with the rule#.
endif::[]
ifeval::["{document-lang}" == "FR"]
[.compliant]#La configuration est en conformité avec la règle#.
endif::[]\n"""
        else:
            rule_file_content += """.{0}
[#nc_{1}, caption="[NC-{2}] - "]
====
{3}
====\n""".format(rule.title, rule.id, "{counter:non-compliance:001}" ,rule.recommendation)

        rule_file_content += "\n"
        with open(f"{str(build_dir / rule.id)}.adoc", "w") as file:
            file.write(rule_file_content)

    def _generate_categories_files(self,
                                    categories: List[Category],
                                    build_dir: Path) -> None:
        for category in categories:
            category_file_content = f"[#{category.category},reftext={category.name}]\n"
            category_file_content += f"== {category.name}\n"
            category_file_content += f"{category.description}\n" if category.description is not None else ""
            category_file_content += "\n\n"

            for rule in category.rules:
                category_file_content += f"include::{str(build_dir / rule.id)}.adoc[]\n"
                self._generate_rule_file(rule, build_dir)

            category_file_content += "\n"
            with open(f"{str(build_dir / category.category)}.adoc", "w") as file:
                file.write(category_file_content)

            self._include_file_in_header(f"{str(build_dir / category.category)}.adoc", build_dir)

    def build_pdf(self,
                    filename: str,
                    output_directory: str,
                    build_dir: str,
                    imagesdir: str = None,
                    pdf_themesdir: str = None,
                    pdf_theme: str = "custom-theme.yml") -> None:

        imagesdir = imagesdir if not imagesdir is None else self._template_dir
        pdf_themesdir = pdf_themesdir if not pdf_themesdir is None else self._template_dir

        if platform.system() == "Windows":
            cmd = ["powershell.exe", f'asciidoctor-pdf -a imagesdir="{imagesdir}/resources/images" -a pdf-themesdir="{pdf_themesdir}/resources/themes" -a pdf-theme="{pdf_theme}" -D "{output_directory}" -o "{filename}.pdf" "{output_directory}/{self._header_file.name}"']
        else:
            cmd = f'asciidoctor-pdf -a imagesdir="{imagesdir}/resources/images" -a pdf-themesdir="{pdf_themesdir}/resources/themes" -a pdf-theme="{pdf_theme}" -D "{output_directory}" -o "{filename}.pdf" "{build_dir}/{self._header_file.name}"'

        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
        process.communicate()

    def generate_pdf(self,
                    filename: str,
                    baseline: Baseline,
                    audited_asset: str,
                    output_directory: Path,
                    imagesdir: str = None,
                    pdf_themesdir: str = None,
                    pdf_theme: str = "custom-theme.yml") -> None:
        if not self._is_asciidoctor_pdf_installed():
            return

        imagesdir = imagesdir if not imagesdir is None else self._template_dir
        pdf_themesdir = pdf_themesdir if not pdf_themesdir is None else self._template_dir

        build_dir = output_directory / "build" / "adoc"
        build_dir.mkdir(parents=True, exist_ok=True)

        report_information = self._initialize_report(filename, baseline.title, audited_asset)
        self._generate_header_file(report_information, build_dir)

        auditee_list_full_name = [x.lstrip().rstrip() for x in report_information["auditee_contact_full_name"].split(';')]
        auditee_list_email = [x.lstrip().rstrip() for x in report_information["auditee_contact_email"].split(';')]
        authors_list_full_name = [x.lstrip().rstrip() for x in report_information["authors_list_full_name"].split(';')]
        authors_list_email = [x.lstrip().rstrip() for x in report_information["authors_list_email"].split(';')]
        self._generate_introduction_file(dict(zip(authors_list_full_name, authors_list_email)), dict(zip(auditee_list_full_name, auditee_list_email)), build_dir)

        self._generate_synthesis_file(baseline.categories, build_dir)
        self._generate_categories_files(baseline.categories, build_dir)

        self.build_pdf(filename, str(output_directory), str(build_dir), imagesdir, pdf_themesdir, pdf_theme)
