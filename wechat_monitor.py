from excel_io import search_aid
from wechat_request import create_session, get_articles_session
import time
import random
import logging


class WechatMonitor:
    def __init__(self, cookies_str, history_path = None, output_path = './wechat_report.xlsx'):
        # list history
        self.aid_history = search_aid(history_path)
        # login
        self.session, self.token = create_session(cookies_str)
        self.output_path = output_path

    def listen_wechats(self, fake_id_l, output_path, itervals = (10, 600)):
        """_summary_

        Args:
            fake_id_l (list): _description_
            output_path (str): _description_
            itervals (tuple, optional): (itervals for fakeid, itervals for times). Defaults to (10, 600).
        """
        while True:
            for fake_id in fake_id_l:
                print(f'\r[{time.asctime( time.localtime(time.time()))}] Listen: {fake_id}', end='')
                article_l = get_articles_session(self.session, fakeid)
                if article_l != 0:
                    self.report(fake_id, article_l)
                # 随机时延
                time.sleep(random.randint(itervals[0] // 2, itervals[0] * 2))
            # 随机时延
            time.sleep(random.randint(itervals[1] // 2, itervals[1] * 2))

    def report(self, fake_id, article_l):
        print('')
        for article in article_l[::-1]:
            logging.warning(f'[New Report!] {fake_id}-{article["time"]}-{article["link"]}')
            self.aid_history.append(article['aid'])
