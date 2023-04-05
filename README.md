# Enterprise-financial-crisis-warning
- data.xlsx内包含了原始数据，来源为CSMAR
- enterprise_1to5_matching.py 用于ST企业与Non-ST企业的1：5匹配，匹配结果见paired_result.xlsx
- annual_report_crawler.py 用于爬取pdf版年报
- stock_bar_comment_crawler.py 用于爬取特定年份的股吧评论，由于请求过多，程序会报异常，正在debug
