# AqSpider

## 说明

抓取真气网的空气质量数据并输出为Excel。

爬取及解密的代码来自：[AKShare/air](https://github.com/akfamily/akshare/tree/master/akshare/air)，在此基础上增加了**命令行参数**及**输出为Excel**的功能以方便使用。

## 使用
进入AqSpider文件夹，运行命令：

```python
py .\aq_spider.py -c 杭州 -s 20220501 -e 20220505 -p hour -l co,no2,o3
```

爬取结果会保存在 `D:\air_quality_results\` 文件夹内，文件名格式为`city_start_end_period.xlsx`，如`杭州_20220501_20220505_hour.xlsx`, 并用电脑默认应用打开。

## 参数说明
| arg | desc |
| --- | --- |
|-c| 城市 |
|-s| 开始时间 |
|-e| 结束时间 |
|-p| 周期，`day` 或 `hour`， `hour` 数据多比较慢，默认为 `day` |
|-l| 需要输出到excel的列，默认全部，即:`'aqi','pm2_5','pm10','co','no2','o3','so2','complexindex','rank','primary_pollutant','temp','humi','windlevel','winddirection','weather'` |
