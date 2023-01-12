# -*- coding: utf-8 -*-
"""
Created on Sun Sep 20 10:13:16 2020

@author: Administrator
"""
#处理数据用
import pandas as pd
import numpy as np
#绘图用
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 用黑体显示中文


class yu(object):
    def __init__(self,This_month_cycle,Last_month_cycle,Last_month,This_month,month_cycle_range):
        #定义下面month_cycle，Last_month ，This_month ，month_cycle_range变量
        #本月周期天数
        self.This_month_cycle = This_month_cycle
        self.Last_month_cycle  = Last_month_cycle 
        #几月份输入格式01，02，03---10，11，12
        self.Last_month = Last_month
        self.This_month = This_month
        #月份周期长度，两位输入
        self.month_cycle_range = month_cycle_range
        
    def _avg(self):
        #input基础数据,DataFrame格式
        Last_month_data = pd.read_excel('【%s%s】月报数据_2866.xlsx' % (self.Last_month,self.month_cycle_range))
        This_month_data = pd.read_excel('【%s%s】月报数据_2866.xlsx' % (self.This_month,self.month_cycle_range))
        
              
        #总体——行驶里程，出班趟次制表
        Total_distance =  round(Last_month_data['总行驶里程km'].sum()/10000 + 0.001,2)
        Total_distance =  np.append(Total_distance,round(This_month_data['总行驶里程km'].sum()/10000 + 0.001,2))
        
        Distance_per_day = round(Last_month_data['日均运行里程km'].sum()/10000 + 0.001,2)
        Distance_per_day = np.append(Distance_per_day,round(This_month_data['日均运行里程km'].sum()/10000 + 0.001,2))
        
        Vehicle_count = Last_month_data['车牌号'].count()
        Vehicle_count = np.append(Vehicle_count,This_month_data['车牌号'].count())
        
        
        Distance_per_day_singlecar = np.round(10000*Distance_per_day/Vehicle_count + 0.001,2)
        
        Total_trip = Last_month_data['周期出班趟次'].sum()
        Total_trip = np.append(Total_trip,This_month_data['周期出班趟次'].sum())
        
        Total_trip_per_day = np.round(Total_trip /self.This_month_cycle,0)
        
        
        Overall_distance_trip_table = pd.DataFrame(  {  '总行驶里程（万公里）' :    Total_distance,   '日均运行里程km'  : Distance_per_day,
                                                      '单车日均行驶里程（公里）':   Distance_per_day_singlecar,'总出班趟次':Total_trip,
                                                      '日均出班趟次':Total_trip_per_day
                                                     }  ,index = ['7月','8月']      )
        
        
        Overall_distance_trip_table.loc['环比增幅']  =  Overall_distance_trip_table.loc['8月'] - Overall_distance_trip_table.loc['7月']
        print('总体——行驶里程，出班趟次表:\n',Overall_distance_trip_table)
        
        #依表绘图
        #月车辆行驶里程做图数据
        x=['7月','8月','环比增幅']
        y1 = Overall_distance_trip_table['总行驶里程（万公里）']
        y2 = Overall_distance_trip_table['日均运行里程km']
        y3 = Overall_distance_trip_table['单车日均行驶里程（公里）']
        bar_width = 0.25
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("%s月车辆行驶里程情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总行驶里程（万公里）', color='steelblue', alpha=1, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均运行里程km', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + 2*bar_width, height=y3, label='单车日均行驶里程（公里）', color='gray', alpha=0.8, width=bar_width)
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y3):
            plt.text(x1 + 2*bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)  
            
        plt.savefig('./%s月车辆行驶里程情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        #月车辆出班趟次情况做图
        x=['7月','8月','环比增幅']
        y1 = Overall_distance_trip_table['总出班趟次']
        y2 = Overall_distance_trip_table['日均出班趟次']
        bar_width = 0.3
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("%s月车辆出班趟次情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总出班趟次', color='orange', alpha=1, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均出班趟次', color='steelblue', alpha=0.5, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        
            
        plt.savefig('./%s月车辆出班趟次情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        #各单位分组下的制表绘图
        #总行驶里程（万公里）	日均行驶里程（万公里）	单车日均行驶里程（公里）	总出班趟次	日均出班趟次	总车数	周期天数
        Last_month_data['总车数'] = 10000
        This_month_data['总车数'] = 10000
        Total_split_Last = Last_month_data.groupby('单位分组')['总行驶里程km','日均运行里程km','总车数','周期出班趟次'].sum()/10000
        Total_split_This = This_month_data.groupby('单位分组')['总行驶里程km','日均运行里程km','总车数','周期出班趟次'].sum()/10000
        Total_split_Last['周期出班趟次'] = Total_split_Last['周期出班趟次']*10000
        Total_split_This['周期出班趟次'] = Total_split_This['周期出班趟次']*10000
        Total_split_Last.loc['其他单位'] = Total_split_Last .loc['专业局'] + Total_split_Last .loc['其他']
        Total_split_This.loc['其他单位'] = Total_split_This .loc['专业局'] + Total_split_This .loc['其他']
        Total_split_Last.drop(['专业局' ,'其他'],axis=0, index=None, columns=None, inplace=True)
        Total_split_This.drop(['专业局' ,'其他'],axis=0, index=None, columns=None, inplace=True)
        Total_split_Last['单车日均行驶里程（公里）'] = 10000*Total_split_Last['日均运行里程km']/Total_split_Last['总车数']
        Total_split_This['单车日均行驶里程（公里）'] = 10000*Total_split_This['日均运行里程km']/Total_split_This['总车数']
        Total_split_Last['日均出班趟次'] = round(Total_split_Last['周期出班趟次']/self.Last_month_cycle)
        Total_split_This['日均出班趟次'] = round(Total_split_This['周期出班趟次']/self.This_month_cycle)
        
        Total_split_Last = Total_split_Last.rename(columns = {'总行驶里程km':'总行驶里程（万公里）','日均运行里程km':'日均行驶里程（万公里）','周期出班趟次':'总出班趟次'})
        Total_split_This = Total_split_This.rename(columns = {'总行驶里程km':'总行驶里程（万公里）','日均运行里程km':'日均行驶里程（万公里）','周期出班趟次':'总出班趟次'})
        
        print(Total_split_Last)
        print(Total_split_This)
        
        #该月车辆运行情况——运输单位
        Transport_split_table = pd.concat([Total_split_Last.loc['运输单位'] ,Total_split_This.loc['运输单位']],axis=1)
        Transport_split_table = pd.DataFrame(Transport_split_table.T)
        Transport_split_table.index = ['7月','8月']
        Transport_split_table.loc['环比增幅'] = round(Transport_split_table.loc['8月'] - Transport_split_table.loc['7月'],2)
        
        #该月车辆运行情况——经营单位
        Manage_split_table = pd.concat([Total_split_Last.loc['经营单位'] ,Total_split_This.loc['经营单位']],axis=1)
        Manage_split_table = pd.DataFrame(Manage_split_table.T)
        Manage_split_table.index = ['7月','8月']
        Manage_split_table.loc['环比增幅'] = round(Manage_split_table.loc['8月'] - Manage_split_table.loc['7月'],2)
        
        
        #该月车辆运行情况——其他单位
        Other_split_table = pd.concat([Total_split_Last.loc['其他单位'] ,Total_split_This.loc['其他单位']],axis=1)
        Other_split_table = pd.DataFrame(Other_split_table.T)
        Other_split_table.index = ['7月','8月']
        Other_split_table.loc['环比增幅'] = round(Other_split_table.loc['8月'] - Other_split_table.loc['7月'],2)
        
        print(Transport_split_table)
        print(Manage_split_table)
        print(Other_split_table)
        
        #运输，经营，其他单位绘图
        
        #【运输】单位组本月车辆行驶里程情况（里程、出班趟次）
        x=['7月','8月','环比增幅']
        y1 = Transport_split_table['总行驶里程（万公里）']
        y2 = Transport_split_table['日均行驶里程（万公里）']
        y3 = Transport_split_table['单车日均行驶里程（公里）']
        bar_width = 0.25
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【运输单位】单位组%s月车辆行驶里程情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总行驶里程（万公里）', color='blue', alpha=0.6, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均行驶里程（万公里）', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + 2*bar_width, height=y3, label='单车日均行驶里程（公里）', color='gray', alpha=0.5, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y3):
            plt.text(x1 + 2*bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        
            
        plt.savefig('./【运输单位】单位组%s月车辆行驶里程情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        x=['7月','8月','环比增幅']
        y1 = Transport_split_table['总出班趟次']
        y2 = Transport_split_table['日均出班趟次']
        bar_width = 0.3
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【运输单位】单位组%s月车辆出班趟次情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总出班趟次', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均出班趟次', color='blue', alpha=0.6, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        
            
        plt.savefig('./【运输单位】单位组%s月车辆出班趟次情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        
        #【经营】单位组本月车辆行驶里程情况（里程、出班趟次）
        x=['7月','8月','环比增幅']
        y1 = Manage_split_table['总行驶里程（万公里）']
        y2 = Manage_split_table['日均行驶里程（万公里）']
        y3 = Manage_split_table['单车日均行驶里程（公里）']
        bar_width = 0.25
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【经营单位】单位组%s月车辆行驶里程情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总行驶里程（万公里）', color='blue', alpha=0.6, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均行驶里程（万公里）', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + 2*bar_width, height=y3, label='单车日均行驶里程（公里）', color='gray', alpha=0.5, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y3):
            plt.text(x1 + 2*bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        
            
        plt.savefig('./【经营单位】单位组%s月车辆行驶里程情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        x=['7月','8月','环比增幅']
        y1 = Manage_split_table['总出班趟次']
        y2 = Manage_split_table['日均出班趟次']
        bar_width = 0.3
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【经营单位】单位组%s月车辆出班趟次情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总出班趟次', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均出班趟次', color='blue', alpha=0.6, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        
            
        plt.savefig('./【经营单位】单位组%s月车辆出班趟次情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        
        #【其他】单位组本月车辆行驶里程情况（里程、出班趟次）
        x=['7月','8月','环比增幅']
        y1 = Other_split_table['总行驶里程（万公里）']
        y2 = Other_split_table['日均行驶里程（万公里）']
        y3 = Other_split_table['单车日均行驶里程（公里）']
        bar_width = 0.25
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【其他单位】单位组%s月车辆行驶里程情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总行驶里程（万公里）', color='blue', alpha=0.6, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均行驶里程（万公里）', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + 2*bar_width, height=y3, label='单车日均行驶里程（公里）', color='gray', alpha=0.5, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y3):
            plt.text(x1 + 2*bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
           
        plt.savefig('./【其他单位】单位组%s月车辆行驶里程情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        x=['7月','8月','环比增幅']
        y1 = Other_split_table['总出班趟次']
        y2 = Other_split_table['日均出班趟次']
        bar_width = 0.3
        
        #设置图形大小
        plt.rcParams['figure.figsize'] = (12.0,8.0) 
        #绘图画板
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        #设置标题
        ax1.set_title("【其他单位】单位组%s月车辆出班趟次情况" % self.This_month,fontsize='18')  
        plt.bar(x=range(len(x)), height=y1, label='总出班趟次', color='orange', alpha=0.8, width=bar_width)
        plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均出班趟次', color='blue', alpha=0.6, width=bar_width)
        
        # 显示图例
        plt.legend()
        plt.xticks(range(len(x)),x)
        # 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
        for x1, yy in enumerate(y1):
            plt.text(x1, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
        for x1, yy in enumerate(y2):
            plt.text(x1 + bar_width, yy + 1, '%.0f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
            
        plt.savefig('./【其他单位】单位组%s月车辆出班趟次情况.png' % self.This_month ,dpi=600,bbox_inches = 'tight')
        
        
        #(二)各区分公司车辆运行情况
        
        Subregion_split_table = pd.DataFrame(This_month_data.groupby('配属单位')['总行驶里程km'].sum())
        Subregion_split_table = pd.merge(Subregion_split_table,Last_month_data.groupby('配属单位')['总行驶里程km'].sum(),on='配属单位',how ='left')
        Subregion_split_table = Subregion_split_table.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
        '北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
        '北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
        '北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
        '北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']]
        Subregion_split_table.index = ['北京市朝阳区分公司','北京市西城区分公司','北京市东城区分公司','北京市海淀区分公司',
        '北京市大兴区分公司','北京市通州区分公司','北京市丰台区分公司',
        '北京市延庆县分公司','北京市房山区分公司','北京市昌平区分公司',
        '北京市门头沟区分公司','北京市密云区分公司','北京市顺义区分公司',
        '北京市石景山区分公司','北京市怀柔区分公司','北京市平谷区分公司']
        Subregion_split_table.columns = ['%s月行驶里程' % self.This_month,'%s月总里程' % self.Last_month]
        Subregion_split_table['差额'] = Subregion_split_table['%s月行驶里程' % self.This_month] -  Subregion_split_table['%s月总里程' % self.Last_month]
        print(Subregion_split_table)
        
        #总行驶里程km<200
        This_month_data.loc[This_month_data['总行驶里程km']<200,'路程<200计数'] = 1
        Subregion_distance200_table = pd.pivot_table(This_month_data,index=['配属单位'],columns=['车型'],values=['路程<200计数'], aggfunc=[np.sum])
        
        Subregion_distance200_table.loc['经营单位合计'] = Subregion_distance200_table.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
        '北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
        '北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
        '北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
        '北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
        Subregion_distance200_table.loc['运输单位合计'] = Subregion_distance200_table.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
        
        Subregion_distance200_table.loc['其他单位合计'] = Subregion_distance200_table.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
        '北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
        '国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司' ]].apply(lambda x: x.sum())
        Subregion_distance200_table.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
        '北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司'],axis = 0,inplace = True)
        Subregion_distance200_table.loc['合计'] = Subregion_distance200_table.loc['运输单位合计'] + Subregion_distance200_table.loc['其他单位合计'] + Subregion_distance200_table.loc['经营单位合计']
        Subregion_distance200_table[ '总计' ] = Subregion_distance200_table. apply ( lambda x: x. sum (), axis = 1 )
        Subregion_distance200_table = Subregion_distance200_table.replace(0,np.nan)
        print(Subregion_distance200_table)
        
        
        #周期出班趟次
        
        This_month_data.loc[This_month_data['周期出班趟次']<10,'趟次<10计数'] = 1
        Subregion_trip10_table = pd.pivot_table(This_month_data,index=['配属单位'],columns=['车型'],values=['趟次<10计数'], aggfunc=[np.sum])
        
        Subregion_trip10_table.loc['经营单位合计'] = Subregion_trip10_table.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
        '北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
        '北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
        '北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
        '北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
        Subregion_trip10_table.loc['运输单位合计'] = Subregion_trip10_table.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
        
        Subregion_trip10_table.loc['其他单位合计'] = Subregion_trip10_table.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
        '北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
        '国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司' ]].apply(lambda x: x.sum())
        Subregion_trip10_table.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
        '北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司'],axis = 0,inplace = True)
        Subregion_trip10_table.loc['合计'] = Subregion_trip10_table.loc['运输单位合计'] + Subregion_trip10_table.loc['其他单位合计'] + Subregion_trip10_table.loc['经营单位合计']
        Subregion_trip10_table[ '总计' ] = Subregion_trip10_table.apply ( lambda x: x. sum (), axis = 1 )
        Subregion_trip10_table =Subregion_trip10_table.replace(0,np.nan)
        print(Subregion_trip10_table)
        
        
        #（五）当月重点单位（行驶公里小于200、出班天数小于10天、运行时长小于30小时）
        This_month_data.loc[((This_month_data['周期出班趟次']<10)&(This_month_data['总行驶里程km']<200)&(This_month_data['总运行时长（点火时长）hh']<30)),'距离200趟次10计数'] = 1
        Subregion_distance200_trip10_time30_table = pd.pivot_table(This_month_data,index=['配属单位'],columns=['车型'],values=['距离200趟次10计数'], aggfunc=[np.sum])
        Subregion_distance200_trip10_time30_table['总计'] =  Subregion_distance200_trip10_time30_table.apply ( lambda x: x. sum (), axis = 1 )
        Subregion_distance200_trip10_time30_table.sort_values(by=['总计'],ascending=False,inplace=True,na_position='first')
        Subregion_distance200_trip10_time30_table = Subregion_distance200_trip10_time30_table.head(5)
        Subregion_distance200_trip10_time30_table.loc[ '合计' ] = Subregion_distance200_trip10_time30_table. apply ( lambda x: x. sum ())
        Subregion_distance200_trip10_time30_table = Subregion_distance200_trip10_time30_table.replace(0,np.nan)
        print(Subregion_distance200_trip10_time30_table)
        
        
        #（六）当月重点车辆（静驶时间过长）序号
        
        Long_stoptime_table = This_month_data[['车牌号','配属单位','名称','吨位','车辆能源类型','总行驶里程km','总运行时长（点火时长）hh'	,'全天静驶时长']]
        Long_stoptime_table['静驶时长占全天运行时长比例'] = Long_stoptime_table['全天静驶时长'] / Long_stoptime_table['总运行时长（点火时长）hh']
        Long_stoptime_table = Long_stoptime_table.loc[Long_stoptime_table['静驶时长占全天运行时长比例']<0.4]
        Long_stoptime_table.sort_values(by=['静驶时长占全天运行时长比例'],ascending=False,inplace=True)
        Long_stoptime_table = Long_stoptime_table.head(10)
        Long_stoptime_table.index = [1,2,3,4,5,6,7,8,9,10]
        Long_stoptime_table['静驶时长占全天运行时长比例'] = pd.DataFrame(Long_stoptime_table['静驶时长占全天运行时长比例']).applymap( lambda x :'%.2f%%' %  (x*100))
        print(Long_stoptime_table)
        
        
        writer = pd.ExcelWriter('需要放入ppt的表(封装后).xlsx')
        Subregion_split_table.to_excel(writer,'各区分公司车辆运行情况')
        Subregion_distance200_table.to_excel(writer,'行驶里程小于200公里车辆')
        Subregion_trip10_table.to_excel(writer,'出班天数小于10天车辆')
        Subregion_distance200_trip10_time30_table.to_excel(writer,'重点单位（里程200出班10时长30）')
        Long_stoptime_table.to_excel(writer,'重点车辆（静驶时间过长）')
        writer.save()
                
 

yu_fengzhuang = yu(31,31,'07','08','01-31')
print(yu_fengzhuang._avg())
