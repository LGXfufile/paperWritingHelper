import pandas as pd
import time
import random

import requests
from fake_useragent import UserAgent


class XHSCrawler:
    def __init__(self):
        self.ua = UserAgent()
        self.session = requests.Session()
        self.proxies = self._init_proxies()
        self.comments_data = []

    def _init_proxies(self):
        # 需替换为实际代理IP列表
        return [
            'http://user:pass@ip:port',
            'https://user:pass@ip:port'
        ]

    def _get_headers(self):
        return {
            'User-Agent': self.ua.random,
            'Referer': 'https://www.xiaohongshu.com/',
            'X-Requested-With': 'XMLHttpRequest'
        }

    def _request_api(self, url, params=None):
        try:
            proxy = random.choice(self.proxies)
            headers = self._get_headers()
            resp = self.session.get(
                url,
                params=params,
                headers=headers,
                proxies={'https': proxy},
                timeout=10
            )
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            print(f'Request failed: {str(e)}')
            return None

    def parse_comments(self, note_id, parent_id=None, level=1):
        api_url = 'https://edith.xiaohongshu.com/api/sns/web/v2/comment/page'
        params = {
            'note_id': note_id,
            'root_comment_id': parent_id or '',
            'num': 20,  # 每页数量
            'cursor': ''
        }

        while True:
            data = self._request_api(api_url, params)
            if not data or not data.get('data'):
                break

            for item in data['data']['comments']:
                self.comments_data.append({
                    'note_id': note_id,
                    'comment_id': item['id'],
                    'parent_id': parent_id or '',
                    'content': item['content'],
                    'like_count': item['like_count'],
                    'time': item['create_time'],
                    'level': level,
                    'user_id': item['user']['user_id']
                })

                # 递归获取子评论
                if item['sub_comment_count'] > 0:
                    self.parse_comments(note_id, item['id'], level + 1)

            if not data['data']['has_more']:
                break

            params['cursor'] = data['data']['cursor']
            time.sleep(random.uniform(1, 3))

    def save_to_excel(self, filename):
        df = pd.DataFrame(self.comments_data)
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f'数据已保存至 {filename}')


if __name__ == '__main__':
    crawler = XHSCrawler()

    # 示例：爬取10篇笔记的评论（需替换实际note_id）
    note_ids = ['65f1a8d9000000001e03d4a1', '65f1b2a3000000001e03d4b2']
    for note_id in note_ids[:10]:
        crawler.parse_comments(note_id)
        time.sleep(random.uniform(3, 5))

    crawler.save_to_excel('xhs_comments.xlsx')
