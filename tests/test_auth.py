#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FeishuAuth 单元测试

测试场景覆盖：
1. 正常获取Access Token
2. Token过期自动刷新
3. 鉴权失败异常处理
4. 无效应用凭证报错

运行方式：
    cd ppt-clone-skill
    python -m pytest tests/test_auth.py -v
    或
    python tests/test_auth.py
"""

import os
import sys
import time
import json
import unittest
from unittest.mock import Mock, patch, MagicMock, call
from datetime import datetime, timedelta
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from feishu.auth import (
    FeishuAuth,
    get_auth,
    FeishuAPIError,
    FeishuAuthError,
    _TOKEN_URL,
    _REFRESH_BUFFER,
)


class TestFeishuAuth(unittest.TestCase):
    """FeishuAuth 单元测试类"""

    @classmethod
    def setUpClass(cls):
        """测试类开始前的初始化"""
        # 保存原始环境变量
        cls.original_env = {
            'FEISHU_APP_ID': os.environ.get('FEISHU_APP_ID'),
            'FEISHU_APP_SECRET': os.environ.get('FEISHU_APP_SECRET'),
        }
    
    def setUp(self):
        """每个测试用例前的初始化"""
        # 设置测试用的环境变量
        os.environ['FEISHU_APP_ID'] = 'test_app_id'
        os.environ['FEISHU_APP_SECRET'] = 'test_app_secret'

    def tearDown(self):
        """每个测试用例后的清理"""
        pass

    @classmethod
    def tearDownClass(cls):
        """测试类结束后的清理"""
        # 恢复原始环境变量
        for key, value in cls.original_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value

    # ==================== 测试用例 1: 正常获取Access Token ====================
    
    @patch('feishu.auth.requests.post')
    def test_01_get_token_success(self, mock_post):
        """测试正常获取Access Token"""
        print("\n[测试1] 正常获取Access Token...")
        
        # 模拟成功的API响应
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'code': 0,
            'msg': 'ok',
            'tenant_access_token': 'test_token_12345',
            'expire': 7200
        }
        mock_post.return_value = mock_response
        
        # 创建认证实例并获取token
        auth = FeishuAuth()
        token = auth.get_token()
        
        # 验证结果
        self.assertEqual(token, 'test_token_12345')
        self.assertEqual(auth._cache.token, 'test_token_12345')
        self.assertIsNotNone(auth._cache.expire_at)
        
        # 验证API调用参数
        mock_post.assert_called_once()
        call_args = mock_post.call_args
        self.assertEqual(call_args[0][0], _TOKEN_URL)
        self.assertEqual(call_args[1]['json'], {
            'app_id': 'test_app_id',
            'app_secret': 'test_app_secret'
        })
        
        print("  ✅ 通过 - Token获取成功")
        print(f"     Token: {token[:20]}...")
        print(f"     过期时间戳: {auth._cache.expire_at}")

    # ==================== 测试用例 2: Token过期自动刷新 ====================
    
    @patch('feishu.auth.requests.post')
    def test_02_token_auto_refresh(self, mock_post):
        """测试Token过期自动刷新"""
        print("\n[测试2] Token过期自动刷新...")
        
        # 第一次调用 - 获取初始token
        mock_response1 = Mock()
        mock_response1.status_code = 200
        mock_response1.json.return_value = {
            'code': 0,
            'msg': 'ok',
            'tenant_access_token': 'initial_token',
            'expire': 7200
        }
        
        # 第二次调用 - 刷新token
        mock_response2 = Mock()
        mock_response2.status_code = 200
        mock_response2.json.return_value = {
            'code': 0,
            'msg': 'ok',
            'tenant_access_token': 'refreshed_token',
            'expire': 7200
        }
        
        mock_post.side_effect = [mock_response1, mock_response2]
        
        # 创建认证实例
        auth = FeishuAuth()
        
        # 第一次获取token
        token1 = auth.get_token()
        self.assertEqual(token1, 'initial_token')
        print(f"  初始Token: {token1}")
        
        # 模拟token过期（将过期时间设为过去）
        auth._cache.expire_at = time.time() - 100
        
        # 第二次获取token，应该自动刷新
        token2 = auth.get_token()
        self.assertEqual(token2, 'refreshed_token')
        self.assertNotEqual(token1, token2)
        
        # 验证API被调用了两次
        self.assertEqual(mock_post.call_count, 2)
        
        print("  ✅ 通过 - Token自动刷新成功")
        print(f"     刷新后Token: {token2}")

    @patch('feishu.auth.requests.post')
    def test_03_token_not_expired_reuse(self, mock_post):
        """测试Token未过期时复用"""
        print("\n[测试3] Token未过期时复用...")
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'code': 0,
            'msg': 'ok',
            'tenant_access_token': 'reusable_token',
            'expire': 7200
        }
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        # 多次获取token
        token1 = auth.get_token()
        token2 = auth.get_token()
        token3 = auth.get_token()
        
        # 验证token相同
        self.assertEqual(token1, token2)
        self.assertEqual(token2, token3)
        
        # 验证API只被调用一次
        mock_post.assert_called_once()
        
        print("  ✅ 通过 - Token复用成功")
        print(f"     获取次数: 3次, API调用: 1次")

    # ==================== 测试用例 3: 鉴权失败异常处理 ====================
    
    @patch('feishu.auth.requests.post')
    def test_04_auth_failure_error_code(self, mock_post):
        """测试API返回错误码时的异常处理"""
        print("\n[测试4] 鉴权失败异常处理 - 错误码...")
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'code': 99991663,
            'msg': 'tenant_access_token not found'
        }
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        with self.assertRaises(FeishuAuthError) as context:
            auth.get_token()
        
        error_msg = str(context.exception)
        self.assertIn('99991663', error_msg)
        self.assertIn('tenant_access_token not found', error_msg)
        
        print("  ✅ 通过 - 正确抛出FeishuAuthError")
        print(f"     错误信息: {error_msg[:60]}...")

    @patch('feishu.auth.requests.post')
    def test_05_http_error_status(self, mock_post):
        """测试HTTP错误状态码处理"""
        print("\n[测试5] HTTP错误状态码处理...")
        
        from requests import HTTPError
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = HTTPError("500 Server Error")
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        with self.assertRaises(FeishuAuthError) as context:
            auth.get_token()
        
        error_msg = str(context.exception)
        self.assertIn('请求 Token 失败', error_msg)
        
        print("  ✅ 通过 - 正确捕获HTTP错误")
        print(f"     错误信息: {error_msg[:60]}...")

    @patch('feishu.auth.requests.post')
    def test_06_network_error(self, mock_post):
        """测试网络错误处理"""
        print("\n[测试6] 网络错误处理...")
        
        from requests.exceptions import RequestException
        mock_post.side_effect = RequestException("Connection timeout")
        
        auth = FeishuAuth()
        
        with self.assertRaises(FeishuAuthError) as context:
            auth.get_token()
        
        error_msg = str(context.exception)
        self.assertIn('Connection timeout', error_msg)
        
        print("  ✅ 通过 - 正确捕获网络异常")
        print(f"     异常信息: {error_msg[:60]}...")

    @patch('feishu.auth.requests.post')
    def test_07_invalid_json_response(self, mock_post):
        """测试无效JSON响应处理"""
        print("\n[测试7] 无效JSON响应处理...")
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.side_effect = json.JSONDecodeError("test", "invalid", 0)
        mock_response.text = "not valid json"
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        with self.assertRaises((FeishuAuthError, json.JSONDecodeError)):
            auth.get_token()
        
        print("  ✅ 通过 - 正确处理无效JSON")

    # ==================== 测试用例 4: 无效应用凭证报错 ====================
    
    def test_08_missing_app_id(self):
        """测试缺少APP_ID时的报错"""
        print("\n[测试8] 无效应用凭证 - 缺少APP_ID...")
        
        # 清除环境变量
        os.environ.pop('FEISHU_APP_ID', None)
        os.environ['FEISHU_APP_SECRET'] = 'test_secret'
        
        with self.assertRaises(FeishuAuthError) as context:
            FeishuAuth()
        
        error_msg = str(context.exception)
        self.assertIn('缺少飞书应用凭证', error_msg)
        
        print("  ✅ 通过 - 正确抛出FeishuAuthError")
        print(f"     错误信息: {error_msg[:50]}...")

    def test_09_missing_app_secret(self):
        """测试缺少APP_SECRET时的报错"""
        print("\n[测试9] 无效应用凭证 - 缺少APP_SECRET...")
        
        # 清除APP_SECRET
        os.environ['FEISHU_APP_ID'] = 'test_app_id'
        os.environ.pop('FEISHU_APP_SECRET', None)
        
        with self.assertRaises(FeishuAuthError) as context:
            FeishuAuth()
        
        error_msg = str(context.exception)
        self.assertIn('缺少飞书应用凭证', error_msg)
        
        print("  ✅ 通过 - 正确抛出FeishuAuthError")
        print(f"     错误信息: {error_msg[:50]}...")

    def test_10_empty_credentials(self):
        """测试空字符串凭证时的报错"""
        print("\n[测试10] 无效应用凭证 - 空字符串...")
        
        os.environ['FEISHU_APP_ID'] = ''
        os.environ['FEISHU_APP_SECRET'] = ''
        
        with self.assertRaises(FeishuAuthError) as context:
            FeishuAuth()
        
        print("  ✅ 通过 - 正确处理空字符串凭证")

    @patch('feishu.auth.requests.post')
    def test_11_invalid_credentials_api_error(self, mock_post):
        """测试API返回无效凭证错误"""
        print("\n[测试11] 无效应用凭证 - API返回错误...")
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'code': 99991661,
            'msg': 'app_id or app_secret is invalid'
        }
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        with self.assertRaises(FeishuAuthError) as context:
            auth.get_token()
        
        error_msg = str(context.exception)
        self.assertIn('99991661', error_msg)
        self.assertIn('app_id or app_secret is invalid', error_msg)
        
        print("  ✅ 通过 - 正确识别无效凭证")
        print(f"     错误码: 99991661")

    # ==================== 测试用例 5: HTTP请求方法测试 ====================
    
    @patch('feishu.auth.requests.get')
    @patch('feishu.auth.requests.post')
    def test_12_get_request_with_auth(self, mock_post, mock_get):
        """测试带鉴权的GET请求"""
        print("\n[测试12] 带鉴权的GET请求...")
        
        # 模拟token获取
        mock_post.return_value = Mock(
            status_code=200,
            json=lambda: {
                'code': 0,
                'msg': 'ok',
                'tenant_access_token': 'test_token',
                'expire': 7200
            }
        )
        
        # 模拟GET请求 - 返回data字段的内容
        mock_get.return_value = Mock(
            status_code=200,
            json=lambda: {'code': 0, 'data': {'test': 'data'}}
        )
        
        auth = FeishuAuth()
        result = auth.get('/test/path', params={'key': 'value'})
        
        # 验证GET请求参数
        mock_get.assert_called_once()
        call_args = mock_get.call_args
        self.assertIn('Authorization', call_args[1]['headers'])
        self.assertIn('Bearer test_token', call_args[1]['headers']['Authorization'])
        
        # auth.get() 返回的是 response.json()['data'] 的内容
        self.assertEqual(result, {'test': 'data'})
        
        print("  ✅ 通过 - GET请求携带正确鉴权头")

    @patch('feishu.auth.requests.post')
    def test_13_post_request_with_auth(self, mock_post):
        """测试带鉴权的POST请求"""
        print("\n[测试13] 带鉴权的POST请求...")
        
        # 模拟token获取
        mock_post.side_effect = [
            Mock(
                status_code=200,
                json=lambda: {
                    'code': 0,
                    'msg': 'ok',
                    'tenant_access_token': 'test_token',
                    'expire': 7200
                }
            ),
            Mock(
                status_code=200,
                json=lambda: {'code': 0, 'data': {'id': '123'}}
            )
        ]
        
        auth = FeishuAuth()
        result = auth.post('/test/path', json={'name': 'test'})
        
        # 验证POST请求参数（第二次调用）
        self.assertEqual(mock_post.call_count, 2)
        call_args = mock_post.call_args
        self.assertIn('Authorization', call_args[1]['headers'])
        self.assertEqual(call_args[1]['json'], {'name': 'test'})
        
        print("  ✅ 通过 - POST请求携带正确鉴权头和JSON数据")

    @patch('feishu.auth.requests.post')
    def test_14_post_multipart_request(self, mock_post):
        """测试multipart/form-data POST请求"""
        print("\n[测试14] Multipart POST请求...")
        
        # 模拟token获取
        mock_post.side_effect = [
            Mock(
                status_code=200,
                json=lambda: {
                    'code': 0,
                    'msg': 'ok',
                    'tenant_access_token': 'test_token',
                    'expire': 7200
                }
            ),
            Mock(
                status_code=200,
                json=lambda: {'code': 0, 'data': {'file_token': 'abc123'}}
            )
        ]
        
        auth = FeishuAuth()
        result = auth.post_multipart(
            '/upload',
            files={'file': ('test.txt', b'content')},
            data={'name': 'test'}
        )
        
        # 验证请求（第二次调用）
        self.assertEqual(mock_post.call_count, 2)
        call_args = mock_post.call_args
        self.assertIn('Authorization', call_args[1]['headers'])
        # multipart时不应设置Content-Type，让requests自动处理
        self.assertNotIn('Content-Type', call_args[1]['headers'])
        
        print("  ✅ 通过 - Multipart请求正确设置头信息")

    # ==================== 测试用例 6: get_auth 辅助函数测试 ====================
    
    @patch('feishu.auth.requests.post')
    def test_15_get_auth_helper(self, mock_post):
        """测试get_auth辅助函数"""
        print("\n[测试15] get_auth辅助函数...")
        
        mock_post.return_value = Mock(
            status_code=200,
            json=lambda: {
                'code': 0,
                'msg': 'ok',
                'tenant_access_token': 'helper_token',
                'expire': 7200
            }
        )
        
        # 调用get_auth获取实例
        auth1 = get_auth()
        auth2 = get_auth()
        
        # 验证返回的是FeishuAuth实例
        self.assertIsInstance(auth1, FeishuAuth)
        self.assertIsInstance(auth2, FeishuAuth)
        
        # get_auth可能实现单例模式，也可能每次都创建新实例
        # 这里只验证返回的是有效实例即可
        
        print("  ✅ 通过 - get_auth返回有效实例")

    # ==================== 测试用例 7: 边界条件测试 ====================
    
    @patch('feishu.auth.requests.post')
    def test_16_token_expiry_edge_case(self, mock_post):
        """测试Token即将过期边界条件"""
        print("\n[测试16] Token即将过期边界条件...")
        
        # 第一次调用返回初始token
        mock_post.side_effect = [
            Mock(
                status_code=200,
                json=lambda: {
                    'code': 0,
                    'msg': 'ok',
                    'tenant_access_token': 'edge_token',
                    'expire': 7200
                }
            ),
            Mock(
                status_code=200,
                json=lambda: {
                    'code': 0,
                    'msg': 'ok',
                    'tenant_access_token': 'new_edge_token',
                    'expire': 7200
                }
            )
        ]
        
        auth = FeishuAuth()
        token1 = auth.get_token()
        
        # 设置token已经过期（当前时间之前），确保触发刷新
        auth._cache.expire_at = time.time() - 10
        
        # 应该触发刷新
        token2 = auth.get_token()
        
        # 验证token已刷新
        self.assertNotEqual(token1, token2)
        self.assertEqual(token2, 'new_edge_token')
        self.assertEqual(mock_post.call_count, 2)
        
        print("  ✅ 通过 - 边界条件处理正确")
        print(f"     旧Token: {token1}")
        print(f"     新Token: {token2}")

    @patch('feishu.auth.requests.post')
    def test_17_empty_response_handling(self, mock_post):
        """测试空响应处理"""
        print("\n[测试17] 空响应处理...")
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {}
        mock_post.return_value = mock_response
        
        auth = FeishuAuth()
        
        with self.assertRaises(FeishuAuthError) as context:
            auth.get_token()
        
        error_msg = str(context.exception)
        self.assertIn('code', error_msg.lower())
        
        print("  ✅ 通过 - 空响应正确处理")

    @patch('feishu.auth.requests.post')
    def test_18_get_headers_method(self, mock_post):
        """测试get_headers方法"""
        print("\n[测试18] get_headers方法...")
        
        mock_post.return_value = Mock(
            status_code=200,
            json=lambda: {
                'code': 0,
                'msg': 'ok',
                'tenant_access_token': 'header_token',
                'expire': 7200
            }
        )
        
        auth = FeishuAuth()
        
        # 默认Content-Type
        headers = auth.get_headers()
        self.assertIn('Authorization', headers)
        self.assertIn('Bearer header_token', headers['Authorization'])
        self.assertEqual(headers['Content-Type'], 'application/json; charset=utf-8')
        
        # 自定义Content-Type
        headers_custom = auth.get_headers(content_type='multipart/form-data')
        self.assertEqual(headers_custom['Content-Type'], 'multipart/form-data')
        
        print("  ✅ 通过 - get_headers返回正确头信息")


class TestReport:
    """测试报告生成器"""
    
    @staticmethod
    def print_summary(result):
        """打印测试摘要"""
        print("\n" + "=" * 70)
        print(" " * 20 + "📊 测试报告摘要")
        print("=" * 70)
        
        total = result.testsRun
        failures = len(result.failures)
        errors = len(result.errors)
        skipped = len(result.skipped)
        passed = total - failures - errors - skipped
        
        print(f"\n  总测试数: {total}")
        print(f"  ✅ 通过: {passed}")
        print(f"  ❌ 失败: {failures}")
        print(f"  💥 错误: {errors}")
        print(f"  ⏭️  跳过: {skipped}")
        
        if failures > 0:
            print("\n" + "-" * 70)
            print("失败用例详情:")
            print("-" * 70)
            for test, traceback in result.failures:
                print(f"\n🔴 {test}")
                print(f"   原因: AssertionError - 预期结果与实际结果不符")
                # 提取关键错误信息
                lines = traceback.strip().split('\n')
                for line in lines[-5:]:
                    if line.strip():
                        print(f"   {line}")
        
        if errors > 0:
            print("\n" + "-" * 70)
            print("错误用例详情:")
            print("-" * 70)
            for test, traceback in result.errors:
                print(f"\n💥 {test}")
                print(f"   原因: 测试执行过程中发生异常")
                # 提取关键错误信息
                lines = traceback.strip().split('\n')
                for line in lines[-5:]:
                    if line.strip():
                        print(f"   {line}")
        
        print("\n" + "=" * 70)
        
        if failures == 0 and errors == 0:
            print("🎉 所有测试通过！代码质量良好。")
        else:
            print(f"⚠️  发现 {failures + errors} 个问题，建议修复后再部署。")
        
        print("=" * 70)


def run_tests():
    """运行所有测试并生成报告"""
    # 创建测试套件
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestFeishuAuth)
    
    # 使用自定义结果类
    class CustomTestResult(unittest.TextTestResult):
        def __init__(self, stream, descriptions, verbosity):
            super().__init__(stream, descriptions, verbosity)
            self.successes = []
        
        def addSuccess(self, test):
            super().addSuccess(test)
            self.successes.append(test)
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2, resultclass=CustomTestResult)
    result = runner.run(suite)
    
    # 打印摘要
    TestReport.print_summary(result)
    
    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    sys.exit(0 if success else 1)
