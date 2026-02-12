from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from dify_plugin.file.file import File

from wps_airscript_client import WPSAirScriptClient


class test_insert_image(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", None)
        file: File = tool_parameters.get("file", None) # type: ignore
        if file is None or file.extension not in ('.jpg','.jpeg','.png','.JPG','.JPEG','.PNG'):
            raise Exception("请上传图片文件！")
        import mimetypes
        mime_type, _ = mimetypes.guess_type(f'img.{file.extension}')
        if mime_type is None:
            mime_type = 'application/octet-stream'  # 默认类型 
        content = None
        try:
            # 至关重要的读取内容的命令！
            # 这一行报错，则说明当前Dify平台的.env全局配置文件里，未设置FILES_URL，导致不同容器间无法共享文件流，该情况会导致灾难性后果，几乎所有文件读写类的第三方插件都会随之失效崩溃，Dify平台管理员请务必注意此项配置要无误！
            content = file.blob
            print('Dify本身里, file.blob读取内容流 -> 成功! ')
        except Exception as e:
            raise Exception("【严重错误】file.blob读取内容流 -> 直接失败！说明当前Dify平台根本不支持后台读取文件流！因此本插件完全无法正常运行！\n可能原因是：\n当前Dify平台的.env全局配置文件里，未设置FILES_URL，导致不同容器间无法共享文件流，该情况会导致灾难性后果，几乎所有文件读写类的第三方插件都会随之失效，并不只是影响本插件了。Dify平台管理员请务必注意此项配置要无误！\n解决办法：\n在.env全局配置文件里，设置：FILES_URL=http://dify服务器的内网IP\n")  
        
        # 添加文件大小校验
        if content is None or len(content) <= 1:
            raise Exception("本次上传的文件内容为空！")
        max_size = 30 * 1024 * 1024  # 单位：MB
        if len(content) > max_size:
            raise Exception(f"本次上传的文件大小超过限制！当前文件大小为 {len(content)/1024/1024:.2f}MB。考虑到Dify平台性能，这里暂时最大只允许 30 MB！")

        if not address:
            raise Exception("请填入单元格地址，例如：A1")

        if address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的单元格地址，例如：A1")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            # 转换为base64
            import base64
            file_base64 = base64.b64encode(content).decode('utf-8')
            file_base64 = f"data:{mime_type};base64,{file_base64}"

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.insert_image(
                address=address,
                image_data=file_base64,
                sheet_name=sheet_name # type: ignore
            )

            if result and result[0]['success']:
                if "插入图片成功" not in result[0]['message']:
                    raise Exception(f"插入图片失败！WPS官方返回错误信息：{result}\n     注：这个命令是AirScript1.0版里的专属命令！如果您当前是默认用的2.0版本的，那么请您再去在线表里手动新建一个1.0版本的脚本！然后单独这一个命令调用时要更换script_id！")  
            else:
                raise Exception(f"插入图片失败！WPS官方返回错误信息：{result}\n 注：这个命令是AirScript1.0版里的专属命令！如果您当前是默认用的2.0版本的，那么请您再去在线表里手动新建一个1.0版本的脚本！然后单独这一个命令调用时要更换script_id！")
        except Exception as e:
            raise Exception(f"插入图片失败！WPS官方返回错误信息：{e}\n 注：这个命令是AirScript1.0版里的专属命令！如果您当前是默认用的2.0版本的，那么请您再去在线表里手动新建一个1.0版本的脚本！然后单独这一个命令调用时要更换script_id！")

        yield self.create_text_message(f"插入图片成功！单元格：{address}，文件名：{file.filename}")
