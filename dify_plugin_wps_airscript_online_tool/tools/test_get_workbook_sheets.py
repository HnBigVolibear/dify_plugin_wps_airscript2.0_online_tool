
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_get_workbook_sheets(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.get_workbook_sheets()

            yield self.create_json_message({"result": result})
            if result:
                yield self.create_text_message(f"{str(result)}\n")
            else:
                yield self.create_text_message("未获取到工作表信息\n")
        except Exception as e:
            raise Exception(f"获取工作表列表失败！WPS官方返回错误信息：{e}")
