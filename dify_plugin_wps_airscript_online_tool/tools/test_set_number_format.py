
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_number_format(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        format_str = tool_parameters.get("format_str", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        if not format_str:
            raise Exception("请输入数字格式字符串")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.set_number_format(address, format_str, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置数字格式失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置数字格式失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置数字格式成功！区域：{address}，格式：{format_str}\n")
