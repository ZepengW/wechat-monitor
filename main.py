import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


import argparse
from wechat_monitor import WechatMonitor


def main(args):
    wm = WechatMonitor(args.config, output_path= args.output)
    #wm.get_newest()
    wm.listen_wechats(itervals=(args.time_accout, args.time_team))


if __name__ == '__main__':
    # 创建一个ArgumentParser对象，设置参数
    parser = argparse.ArgumentParser(description='My program')

    # 添加参数
    parser.add_argument('--config', '-f', type=str, help='config path', default='./tmp/cfg.yml')
    parser.add_argument('--output', '-o', type=str, help='ourput path', default='./tmp/articles.xlsx')
    parser.add_argument('--time_accout', '-ta', type=int, help='每个公众号的访问间隔，默认60s', default=60)
    parser.add_argument('--time_team', '-tt', type=int, help='每组公众号的访问间隔，默认3600s', default=3600)


    # 解析参数
    args = parser.parse_args()

    # 调用主程序入口函数
    main(args)
