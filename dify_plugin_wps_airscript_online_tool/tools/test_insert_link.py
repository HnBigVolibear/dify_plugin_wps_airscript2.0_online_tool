from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient


class test_insert_link(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        url = tool_parameters.get("url", "").strip()
        text = tool_parameters.get("text", "").strip()
        sheet_name = tool_parameters.get("sheet_name", None)

        if not text:
            text = url

        if not address:
            raise Exception("请填入单元格地址，例如：A1")

        if address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的单元格地址，例如：A1")

        if not url or not (url.startswith("http://") or url.startswith("https://")):
            raise Exception("请填入链接网址")
        if len(url) > 99999:
            raise Exception("链接网址过长，最长允许99999个字符")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.insert_link(
                address=address,
                text=text,
                url=url,
                sheet_name=sheet_name # type: ignore
            )

            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"插入链接失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"插入链接失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"插入链接成功！单元格：{address}，链接：{url}")
