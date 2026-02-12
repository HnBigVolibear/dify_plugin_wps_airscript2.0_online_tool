
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_copy_paste_range(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        source_address = tool_parameters.get("source_address", "").strip()
        target_address = tool_parameters.get("target_address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not source_address or source_address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的源区域地址，例如：A1:C3")

        if not target_address or target_address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的目标起点单元格地址，例如：E1")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.copy_paste_range(source_address, target_address, sheet_name)
            if result and result[0]['success']:
                yield self.create_text_message(f"复制粘贴成功！\n")
            else:
                # raise Exception(f"复制粘贴失败！WPS官方返回错误信息：{result}")
                yield self.create_text_message(f"复制粘贴失败！\n")
        except Exception as e:
            raise Exception(f"复制粘贴失败！WPS官方返回错误信息：{e}")
