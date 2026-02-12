
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_get_filtered_data(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_name = tool_parameters.get("sheet_name", None)

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.get_filtered_data(sheet_name) # type: ignore

            if result and result[0]['success']:
                data = result[0]['data']
                row_count = result[0]['rowCount']
                column_count = result[0]['colCount']
            else:
                raise Exception(f"获取筛选后数据失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"获取筛选后数据失败！WPS官方返回错误信息：{e}")

        yield self.create_json_message({
            "result": data,
            "rowCount": row_count,
            "colCount": column_count
        })

        if not data or row_count == 0:
            yield self.create_text_message("当前没有筛选条件或筛选后无数据！")
        else:
            yield self.create_text_message(f"获取筛选后数据成功！\n共 {row_count} 行 {column_count} 列数据。\n具体数据如下：\n{data}")
