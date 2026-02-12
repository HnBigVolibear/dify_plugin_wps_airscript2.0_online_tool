
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_worksheet_exists(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_name = tool_parameters.get("sheet_name", "")

        if not sheet_name:
            raise Exception("请输入工作表名称")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.worksheet_exists(sheet_name)

            yield self.create_json_message({"result": result})
            yield self.create_text_message(f"工作表 '{sheet_name}' {'存在' if result else '不存在'}\n")
        except Exception as e:
            raise Exception(f"检查工作表是否存在失败！WPS官方返回错误信息：{e}")
