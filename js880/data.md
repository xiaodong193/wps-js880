J1公式是：=k("$$.superPivot",A1:H40,"f3,f2","f6","count(),sum('f4'),textjoin('f4','+')")

数据：A1:H40，带有表头。
数据：
ID	产品	国家	数量	价格	年	月	日
1	Product1	中国	1	1	2021	10	10
2	Product2	usa	19	5	2023	4	5
3	Product2	英国	19	5	2022	6	28
4	Product2	美国	15	5	2024	5	1
5	Product1	中国	11	1	2024	11	15
6	Product2	德国	18	5	2023	2	18
7	Product2	英国	11	5	2023	6	16
8	Product2	美国	11	5	2023	6	21
9	Product1	中国	13	1	2022	7	18
10	Product1	德国	18	1	2021	11	13
11	Product2	英国	11	5	2022	10	7
12	Product1	美国	16	1	2023	5	15
13	Product1	中国	11	1	2022	9	2
14	Product1	德国	12	1	2024	11	25
15	Product1	英国	16	1	2023	8	22
16	Product1	美国	16	1	2022	4	22
17	Product2	中国	15	5	2023	9	17
18	Product1	德国	13	1	2023	4	28
19	Product1	英国	18	1	2021	11	20
20	Product1	美国	12	1	2023	7	13
21	Product1	中国	18	1	2021	1	23
22	Product1	德国	16	1	2023	2	18
23	Product2	英国	15	5	2022	9	1
24	Product1	美国	12	1	2023	2	21
25	Product1	中国	17	1	2023	11	2
26	Product1	德国	15	1	2023	11	16
27	Product1	英国	15	1	2024	10	23
28	Product2	美国	13	5	2023	1	26
23	Product2	test	15	5	2022	9	1
24	Product1	test	12	1	2023	2	21
25	Product1	test	17	1	2023	11	2
26	Product1	test	15	1	2023	11	16
27	Product1	test	15	1	2024	10	23
28	Product2	test	13	5	2023	1	26
24	Product1	test	12	1	2023	2	21
25	Product1	test	17	1	2023	11	2
26	Product1	test	15	1	2023	11	16
27	Product1	test	15	1	2024	10	23
28	Product2	test	13	5	2023	1	26
结果：
年	年	2021	2021	2021	2022	2022	2022	2023	2023	2023	2024	2024	2024
国家	产品	计数	求和	多项合并	计数	求和	多项合并	计数	求和	多项合并	计数	求和	多项合并
德国	Product1	1	18	18				3	44	13+16+15	1	12	12
德国	Product2							1	18	18			
美国	Product1				1	16	16	3	40	16+12+12			
美国	Product2							2	24	11+13	1	15	15
英国	Product1	1	18	18				1	16	16	1	15	15
英国	Product2				3	45	19+11+15	1	11	11			
中国	Product1	2	19	1+18	2	24	13+11	1	17	17	1	11	11
中国	Product2							1	15	15			
test	Product1							6	88	12+17+15+12+17+15	2	30	15+15
test	Product2				1	15	15	2	26	13+13			
usa	Product2							1	19	19			

期望结果：当前多级表头，如下：
	年	2021	2021	2021	2022	2022	2022	2023	2023	2023	2024	2024	2024
国家	产品	计数	求和	多项合并	计数	求和	多项合并	计数	求和	多项合并	计数	求和	多项合并
德国	Product1	1	18	18				3	44	13+16+15	1	12	12
德国	Product2							1	18	18			
美国	Product1				1	16	16	3	40	16+12+12			
美国	Product2							2	24	11+13	1	15	15
英国	Product1	1	18	18				1	16	16	1	15	15
英国	Product2				3	45	19+11+15	1	11	11			
中国	Product1	2	19	1+18	2	24	13+11	1	17	17	1	11	11
中国	Product2							1	15	15			
test	Product1							6	88	12+17+15+12+17+15	2	30	15+15
test	Product2				1	15	15	2	26	13+13			
usa	Product2							1	19	19			
