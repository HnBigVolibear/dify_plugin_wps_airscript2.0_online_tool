
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_cell_formula(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        formula = tool_parameters.get("formula", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的单元格地址，例如：A1、或G1:G30")

        if not formula:
            raise Exception("请输入要设置的公式")
        elif formula[0] != "=":
            formula = "=" + formula

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.set_cell_formula(address, formula, sheet_name)
            if result and result[0]['success']: 
                pass
            else:
                raise Exception(f"设置单元格公式失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置单元格公式失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message("设置单元格公式成功！\n")
