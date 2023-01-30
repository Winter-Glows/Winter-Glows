import os
import json
import base64
import requests
import datetime

SERVER_URL = "https://zt.precisiongenes.com.cn/api/v1"

""" 使用者换成自己的账号，这边是示例"""
USER_INFO = {
    "username": "18256069387",
    "password": "18256069387"
}

URL_KEY = {
    "token_auth": "/token-auth/",
    "order": "/order/",
    "reportfile": "/reportfile/",
    "report": "/report/",
    "mini_user": "/mini_user/",
    "sample": "/sample/",
    "user": "/user/",
    "orderproduct": "/orderproduct/",
    "product": "/product/",
    "pedigree": "/pedigree/",
    "pedigreeperson": "/pedigreeperson/"
}

class Base:
    """
    中台 SDK

    具体支持的 api 请参照 https://zt.precisiongenes.com.cn/api/v1/zkzt-swagger/
    可以通过 apifox 软件导入本链接，来查看所有端点使用方式，并支持生成接口代码
    注意⚠️：中台请求需要授权，在 requests headers 里面添加键值对  Authorization: JWT <token>
    """

    def __init__(self):
        """
        初始化获取 token
        """
        self._token = ""
        self.user_info = None
        self._get_token()
        self.get_user_url()

    def do_request(self, url_key: str, method: str, params: dict = None, data: dict = None,
                   need_auth: bool = True, files=None) -> json:
        """
        向中台发起请求

        :param url_key: config settings 里面设置请求路径
        :param method: 请求方法，可以支持 [GET POST PUT PATCH DELETE], 暂时只支持 get、post 其他方法请自行扩展
        :param params: 请求链接里面的字典参数
        :param data: 请求体字典参数
        :param need_auth: 是否需要认证, 默认需要
        :param files: 需要上传的文件
        :return: 200 状态下的 r.json()
        """
        assert URL_KEY.get(url_key, None) and method, "url 和 method 不能为空"
        assert method in ["get", "post", "patch"], "非法请求方式！"
        headers = {} if not need_auth else {
            "Authorization": "JWT " + self._token,
        }
        url = SERVER_URL + URL_KEY.get(url_key)
        if method == "get":
            r = requests.get(url=url, params=params, headers=headers)
        elif method == "patch":
            url = url + '{}/'.format(str(params.get('id')))
            r = requests.patch(url=url, data=data, headers=headers)
        else:
            r = requests.post(url=url, params=params, data=data, headers=headers, files=files)
        # print(r.status_code, r.text)
        if r.status_code in [401, 403]:
            # 授权失效的话重新获取
            self._get_token()
            self.do_request(url_key=url_key, method=method, params=params)
        # 其他状态码可以自行往下判断
        if r.status_code < 300:
            return r.json()
        else:
            return None

    def get_user_url(self):
        # print(self._token.split("."))
        base_user_info = self._token.split(".")[1]
        if len(base_user_info) % 4:
            base_user_info = base_user_info + "=" * (len(base_user_info) % 4)
        assert type(base_user_info) == str, "token 无信息，请重新运行脚本！"
        # print(user_base_info)
        user_json_info = json.loads(base64.b64decode(base_user_info))
        # print(user_json_info)
        mini_user_url = SERVER_URL + URL_KEY.get("mini_user", None) + str(user_json_info.get("user_id")) + "/"
        user_url = SERVER_URL + URL_KEY.get("user", None) + str(user_json_info.get("user_id")) + "/"
        self.user_info = user_json_info
        self.user_info['user_url'] = user_url
        # print(self.user_info)

    def _get_token(self):
        """
        获取 token, 类初始化的时候会调用，无需手动调用

        :return:
        """
        auth = self.do_request(url_key="token_auth", method="post", need_auth=False, data=USER_INFO)
        if auth:
            self._token = auth.get("token")
        # print(self._token)

    def get_order(self):
        """
        获取订单接口，示例，请删除并继承基类重载
        :return:
        """
        my_orders = self.do_request(url_key="order", method="get")
        # print(my_orders)


if __name__ == '__main__':
    base_class = Base()
    # base_class.get_order()
    base_class.get_user_url()
