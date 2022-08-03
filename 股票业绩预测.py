import execjs
import pandas as pd
import requests
from loguru import logger
from scrapy import Selector


class Forecast(object):
    """
    股票业绩预测
    useage:
        1.在stocks.xlsx文件 股票名称列 添加你想要预测的个股 如:贵州茅台 保存退出
        2.运行 股票业绩预测.py 执行完毕会将所有iwencai采集结果写入stocks.xlsx 并按年复合增长率从大到小进行排序
        3.选择 年复合增长率大于30%品种 优先选择具备行业垄断地位的龙头企业
        4.温馨提示: 投资有风险 入市需谨慎
        5.欢迎star、交流！
    """
    def __init__(self):
        self.session = requests.session()
        self.stocks_path = 'stocks.xlsx'

        js_code_path = r'get_cookie_v.js'
        with open(js_code_path, 'r', encoding='utf-8') as f:
            js_code = f.read()
        self.js_code_compile = execjs.compile(js_code)

    def query_one(self, stock_name):
        v = self.js_code_compile.call('getCookieV')
        logger.info(v)

        w = f'{stock_name}业绩预测'
        url = 'http://www.iwencai.com/stockpick/search?typed=0&preParams=&ts=1&f=1&qs=result_original' + \
              f'&selfsectsn=&querytype=stock&searchfilter=&tid=stockpick&w={w}'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)' +
                          'Chrome/92.0.4515.107 Safari/537.36',
            'Referer': 'http://www.iwencai.com/stockpick/search?typed=1&preParams=&ts=1&f=1' +
                       '&qs=result_rewrite&selfsectsn=&querytype=stock&searchfilter=' +
                       '&tid=stockpick&w=%E5%8D%93%E8%83%9C%E5%BE%AETTM&queryarea=',
            'Cookie': v
        }
        response = self.session.get(url, headers=headers, timeout=3)

        response = Selector(response=response)
        years = response.xpath(
            '//table[@class="upright_table"]//th[@class="up_th marge_th"]/div[@class="em"]/text()').re(r'[\d\.]+')
        profits = response.xpath(
            '//table[@class="upright_table"]//div[@class="em alignRight alignRight"]/text()').getall()
        profits = [profit.replace('亿', '').replace(',', '') for profit in profits]

        year_profit = list(zip(years, profits))
        ret = self.two_year(profits[0], profits[2])

        return year_profit, ret

    def query_all(self):
        df = pd.read_excel(self.stocks_path)

        stock_names = df['股票名称'].values
        for stock_name in stock_names:
            try:
                year_profit, ret = self.query_one(stock_name)
                logger.info(f'{stock_name} {year_profit} {ret}')
                self.update_excel(df, stock_name, year_profit, ret)
            except Exception as e:
                logger.info(f'query_all {e}')

        # sort df
        df.sort_values(by='机构预测结果', inplace=True, ascending=False)

        # save result
        df.to_excel(self.stocks_path, index=False)

    def update_excel(self, df, stock_name, year_profit, ret):
        df.loc[df['股票名称'] == stock_name, '机构预测详情'] = f'{year_profit}'
        df.loc[df['股票名称'] == stock_name, '机构预测结果'] = ret

    def two_year(self, three_year, first_year):
        try:
            ret = (float(three_year) / float(first_year)) ** (1 / 2)
            ret = round((ret - 1) * 100, 2)
        except Exception as e:
            logger.info(f'two_year exception {e}')
            ret = -1

        return ret


if __name__ == '__main__':
    forecast = Forecast()
    forecast.query_all()
