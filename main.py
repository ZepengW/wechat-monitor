import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


import argparse
from wechat_monitor import WechatMonitor


def main(args):
    wm = WechatMonitor(args.config, output_path= args.output)
    wm.check_ready()
    wm.get_newest()
    wm.listen_wechats()


if __name__ == '__main__':
    # 创建一个ArgumentParser对象，设置参数
    parser = argparse.ArgumentParser(description='My program')

    # 添加参数
    parser.add_argument('--config', '-f', type=str, help='config path', default='./tmp/cfg.yml')
    parser.add_argument('--output', '-o', type=str, help='ourput path', default='./tmp/articles.xlsx')

    # 解析参数
    args = parser.parse_args()

    # 调用主程序入口函数
    main(args)
