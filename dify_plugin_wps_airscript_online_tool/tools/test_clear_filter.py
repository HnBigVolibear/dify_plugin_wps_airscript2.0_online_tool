
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_clear_filter(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_name = tool_parameters.get("sheet_name", "").strip()

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.clear_filter(sheet_name if sheet_name else None) # type: ignore

            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"清除筛选失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"清除筛选失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message("清除筛选成功！")
