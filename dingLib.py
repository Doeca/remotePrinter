from alibabacloud_dingtalk.oauth2_1_0.client import Client as dingtalkoauth2_1_0Client
from alibabacloud_dingtalk.oauth2_1_0 import models as dingtalkoauth_2__1__0_models
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_tea_util.client import Client as UtilClient
from alibabacloud_dingtalk.workflow_1_0.client import Client as dingtalkworkflow_1_0Client
from alibabacloud_dingtalk.workflow_1_0 import models as dingtalkworkflow__1__0_models
from alibabacloud_tea_util import models as util_models

import time
import requests
from logsys import logger
import db

def class_to_dict(obj):
    if isinstance(obj, dict):
        return {k: class_to_dict(v) for k, v in obj.items()}
    elif hasattr(obj, "_ast"):
        return class_to_dict(obj._ast())
    elif hasattr(obj, "__iter__") and not isinstance(obj, str):
        return [class_to_dict(v) for v in obj]
    elif hasattr(obj, "__dict__"):
        return {k: class_to_dict(v) for k, v in obj.__dict__.items() if not k.startswith('_')}
    else:
        return obj


def getToken() -> str:
    config = open_api_models.Config()
    config.protocol = 'https'
    config.region_id = 'central'
    client = dingtalkoauth2_1_0Client(config)
    get_access_token_request = dingtalkoauth_2__1__0_models.GetAccessTokenRequest(
        app_key='dingv0nxcrvyxbxhevny',
        app_secret='Bw6rLH04ro40KTyf8EF_Z9mCUCV9b_i-z077Na9qlfN8uBoRuyA3YZxb9HA9sFpL'
    )
    try:
        res = client.get_access_token(get_access_token_request)
        return res.body.access_token
    except Exception as err:
        if not UtilClient.empty(err.code) and not UtilClient.empty(err.message):
            # err 中含有 code 和 message 属性，可帮助开发定位问题
            logger.error(err.code)
            logger.error(err.message)


def getInstances(processcode: str, p_statuses: list = []):

    # 能否本地维护一个缓存表，当每页的数据的出现本地缓存表中的数据时，就停止请求，返回本地缓存，减少请求次数
    p_next_token = 1
    res_list = {"list": []}
    while p_next_token != None:
        config = open_api_models.Config()
        config.protocol = 'https'
        config.region_id = 'central'
        client = dingtalkworkflow_1_0Client(config)
        headers = dingtalkworkflow__1__0_models.ListProcessInstanceIdsHeaders()
        headers.x_acs_dingtalk_access_token = getToken()
        requestbody = dingtalkworkflow__1__0_models.ListProcessInstanceIdsRequest(
            process_code=processcode,
            start_time=(int(time.time()) - 86400 * 5)*1000,
            next_token=p_next_token,
            max_results=20,
            statuses=['COMPLETED'] if p_statuses == [] else p_statuses
        )
        try:
            res = client.list_process_instance_ids_with_options(
                requestbody, headers, util_models.RuntimeOptions())
            p_next_token = res.body.result.next_token
            # 遍历当前页的数据，如果有缓存的数据，就停止请求，返回本地缓存
            # 没有数据，就获取表单详情，加入本地缓存表中，同时清空90天以前的表id
            for i in res.body.result.list:
                if not db.check_ids_cache_exists(processcode, i):
                    instance_data = class_to_dict(getDetail(i))
                    logger.debug(f"获取表单详情_{i}，加入本地缓存")
                    timestamp = int(time.mktime(time.strptime(
                        instance_data['create_time'], "%Y-%m-%dT%H:%MZ")))
                    db.add_ids_cache(processcode, i, timestamp)
        except Exception as err:
            logger.error(err)
            if not UtilClient.empty(err.code) and not UtilClient.empty(err.message):
                # err 中含有 code 和 message 属性，可帮助开发定位问题
                pass
    res_list['list'] = get_cache_within_90days(processcode)
    return res_list


def get_cache_within_90days(processcode: str):
    p_info = db.get_ids_cache_by_process(processcode)
    if not p_info:
        return []
    threshold = int(time.time()) - 86400 * 90
    # 删除过90天的缓存
    db.delete_old_ids_cache(processcode, threshold)
    # 返回删除后的数据
    p_info = db.get_ids_cache_by_process(processcode)
    return list(p_info.keys()) if p_info else []


def getDetail(processID: str):
    config = open_api_models.Config()
    config.protocol = 'https'
    config.region_id = 'central'
    client = dingtalkworkflow_1_0Client(config)
    get_process_instance_headers = dingtalkworkflow__1__0_models.GetProcessInstanceHeaders()
    get_process_instance_headers.x_acs_dingtalk_access_token = getToken()
    get_process_instance_request = dingtalkworkflow__1__0_models.GetProcessInstanceRequest(
        process_instance_id=processID
    )
    try:
        res = client.get_process_instance_with_options(
            get_process_instance_request, get_process_instance_headers, util_models.RuntimeOptions())
        return res.body.result
    except Exception as err:
        if not UtilClient.empty(err.code) and not UtilClient.empty(err.message):
            # err 中含有 code 和 message 属性，可帮助开发定位问题
            pass


def getAttachment(processID: str, fileid: str):
    config = open_api_models.Config()
    config.protocol = 'https'
    config.region_id = 'central'
    client = dingtalkworkflow_1_0Client(config)
    grant_process_instance_for_download_file_headers = dingtalkworkflow__1__0_models.GrantProcessInstanceForDownloadFileHeaders()
    grant_process_instance_for_download_file_headers.x_acs_dingtalk_access_token = getToken()
    grant_process_instance_for_download_file_request = dingtalkworkflow__1__0_models.GrantProcessInstanceForDownloadFileRequest(
        process_instance_id=processID,
        file_id=fileid
    )
    try:
        res = client.grant_process_instance_for_download_file_with_options(
            grant_process_instance_for_download_file_request, grant_process_instance_for_download_file_headers, util_models.RuntimeOptions())
        return res.body.result
    except Exception as err:
        if not UtilClient.empty(err.code) and not UtilClient.empty(err.message):
            # err 中含有 code 和 message 属性，可帮助开发定位问题
            pass


def getUserName(userID):
    try:
        resp = requests.post(f"https://oapi.dingtalk.com/topapi/v2/user/get?access_token={getToken()}", data={
            'userid': userID
        })
        res = resp.json()
        if res.get('errcode', -1) != 0:
            logger.warning(f"获取用户信息失败: {res.get('errmsg', 'Unknown')}")
            return '【不存在】'
        result = res.get('result', {})
        return result.get('name', '【未知用户】')
    except Exception as e:
        logger.error(f"获取用户名称失败: {e}")
        return '【错误】'
# res = getDetail("Yzh3xDSBREyl3B1QtRJaRw03401681633674")
# a = 'https://static.dingtalk.com/media/lQDPM5esycx2MzrNBaDNBDiwlXedu-SDBMsEM-l4LgDQAA_1080_1440.jpg'

# getUserName('2110673635943098')
