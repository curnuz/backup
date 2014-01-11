#!/usr/bin/python
#coding=utf-8

import xlrd

from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, Float, String, DateTime
from sqlalchemy.orm import sessionmaker
from datetime import datetime

import sys
import rpy2.robjects as ro

#######################################
## 初始化配置信息
#######################################
INIT_ASSET = 106831.41
INIT_OPEN_DATE = '2013-09-01'
DB_USER = 'lifesurge'
DB_PASSWORD = '111111'
DB_NAME = 'mozart'
########################################

engine = create_engine('postgresql://%s:%s@localhost:5432/%s' \
	% (DB_USER, DB_PASSWORD, DB_NAME), echo=False)
Base = declarative_base()
Session = sessionmaker(bind=engine)
session = Session()

class TradeLog(Base):
	__tablename__ = 'trade_logs'
	id = Column(Integer, primary_key=True, autoincrement=True)
	asset = Column(Float, nullable=False)
	profit = Column(Float, nullable=False)
	open_date = Column(String, nullable=False)
	close_date = Column(String, nullable=False)
	created = Column(DateTime, nullable=False)
	
	def __init__(self):
		self.asset = 0
		self.profit = 0
		self.open_date = '0'
		self.close_date = '0'
		self.created = datetime.now()
	
def get(set):
	ts = session.query(TradeLog).order_by(TradeLog.created).all()
	for t in ts:
		set.append(t)
		
def process_xls(openclose, name):
	bk = xlrd.open_workbook(name)
	sh = bk.sheet_by_index(3)
	nrows = sh.nrows
	#print nrows

	start = False
	for i in range(nrows - 1):
		if sh.cell_value(i, 0) == u"交易日期":
			start = True
		elif start:
			fee = sh.cell_value(i, 10)
			profit = sh.cell_value(i, 11)
			date = sh.cell_value(i, 12)

			current = TradeLog()
			if len(openclose) > 0: 
				current = openclose[-1]

			if type(profit) is unicode: ## 开仓
				if openclose[-1].close_date == '0':
					current.profit -= fee
					current.asset -= fee
					session.merge(current)
				else:
					current = TradeLog()
					current.profit -= fee
					current.asset = openclose[-1].asset
					current.asset += current.profit
					current.open_date = date
					openclose.append(current)
					session.add(current)
				session.commit()
			else: ## 平仓
				current.profit -= fee
				current.profit += profit
				current.asset += profit - fee
				current.close_date = date
				session.merge(current)
				session.commit()
				
def report(openclose):
	if openclose[-1].close_date == '0':
		openclose = openclose[:-1]
	
	w = 0
	wt = 0
	l = 0
	lt = 0
	max_lt = 0
	c_lt = 0
	max_l = 0
	c_l = 0
	
	max_b_p = 0.0
	max_b_p_a = 0
	
	for i in openclose:
		if i.profit > 0:
			w += i.profit
			wt += 1
			c_lt = 0
			c_l = 0
			max_b_p_a = i.asset
		else:
			c_lt += 1
			if c_lt > max_lt:
				max_lt = c_lt
			c_l += i.profit
			if c_l < max_l:
				max_l = c_l
			l += i.profit
			lt += 1

			if max_b_p_a > 0:
				b_p = 1.0*max_l/max_b_p_a
				if b_p < max_b_p:
					max_b_p = b_p
					

	print '期初日期:', (openclose[0].open_date),\
		'\t期末日期:', (openclose[-1].open_date)
	print '期初权益: %.2f' % (106831.41),\
		'\t期末权益: %.2f' % (openclose[-1].asset)
	print '净盈利金额: %.2f' % (openclose[-1].asset - openclose[0].asset),\
		'\t净盈利比: %.2f%%' % (100.0*(openclose[-1].asset - \
		openclose[0].asset)/106831.41)

	print '总交易次数:', (wt + lt),\
		'\t总盈利次数:', wt
		
	print '总盈利金额: %.2f' % w,'\t总亏损金额: %.2f' % l
	print '平均每笔盈利: %.2f' % (1.0*w/wt),\
		'\t平均每笔亏损: %.2f' % (1.0*l/lt)
	print '\t胜率: %.2f%%' % (100.0*wt/(wt+lt)),\
		'\t盈亏比: %.2f' % (1.0*w/wt/(-l/lt))
	print '最大回撤比: %.2f%%' % (100.0*max_b_p),\
		'\t最大连续亏损次数: %d' % max_lt



if __name__ == '__main__':
	usage = '''
	#初始化
	./mozart.py init
	#查看报告
	./mozart.py report
	#处理保证金监控中心月度报告
	./mozart.py *.xls
	'''
	if len(sys.argv) != 2:
		print usage
		exit()
	if sys.argv[1] == 'init':
		Base.metadata.create_all(engine)
	
		t = TradeLog()
		t.asset = INIT_ASSET
		t.profit = 0
		t.open_date = INIT_OPEN_DATE
		t.close_date = '0'
		session.add(t)
		session.commit()
		print '初始化成功'

	elif sys.argv[1] == 'report':
		openclose = []
		get(openclose)
		report(openclose)
		print '获取报告成功'
	elif sys.argv[1] == 'value':
		openclose = []
		get(openclose)
		if openclose[-1].close_date == '0':
			openclose = openclose[:-1]
		x = [0]
		y = [openclose[0].asset-openclose[0].profit]
		max_asset = 0
		min_asset = 0
		j = 1
		for i in openclose:
			#x.append(i.close_date)
			x.append(j)
			j += 1
			y.append(i.asset)
			if i.asset > max_asset:
				max_asset = i.asset
			if i.asset < min_asset or min_asset == 0:
				min_asset = i.asset
		r = ro.r
		r.png('value.png', width=800, height=500)
		r.plot(x, y, type="l", col="black", xlab="日期",\
			ylab="权益", main="ArchDark期货程序化交易系统",\
			lwd=3, las=1)
		r.legend(x=0,y=max_asset, legend=r.c("Close: %.2f  Open: %.2f"%(openclose[-1].asset, (openclose[0].asset - openclose[0].profit)), "Max: %.2f  Min: %.2f" % (max_asset, min_asset)))
		#r.grid( ny=8,lwd=1,lty=2,col="blue")
		r.abline(h=openclose[0].asset-openclose[0].profit, col="red", lty=2)
		#r.dev.off()
	elif sys.argv[1] == 'bar':
		openclose = []
		get(openclose)
		if openclose[-1].close_date == '0':
			openclose = openclose[:-1]

		x = []
		y = []
		max_win = 0
		max_loss = 0
		avg_win = 0
		avg_loss = 0
		win_t = 0
		loss_t = 0
		j = 0
		for i in openclose:
			#x.append(i.close_date)
			x.append(j)
			j += 1
			y.append(i.profit)
			if i.profit > max_win:
				max_win = i.profit
			if i.profit < max_loss:
				max_loss = i.profit
			if i.profit > 0:
				avg_win += i.profit
				win_t += 1
			else:
				avg_loss += i.profit
				loss_t += 1
		r = ro.r
		r.png('bar.png', width=800, height=500)
		r.plot(x, y, type="h", col="black", xlab="日期",\
			ylab="单次盈亏", main="ArchDark期货程序化交易系统",\
			lwd=3, las=1)
		r.legend(x=0,y=max_win,legend=r.c("MaxWin: %.2f  AvgWin: %.2f" % (max_win, 1.0*avg_win/win_t), \
		"MaxLoss: %.2f  AvgLoss: %.2f" % (max_loss,  1.0*avg_loss/loss_t)))
		#r.abline(h=0, col="red", lty=1)

	else:
		openclose = []
		get(openclose)
		process_xls(openclose, sys.argv[1])
	
		print '导入数据: %s 成功', sys.argv[1]
