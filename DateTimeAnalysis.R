# This R Script reads an xlsx file containing raw card swiping data from bjfu.sunnysport.org.cn,
# and mainly process datetime data. The analysis items include(item - output):
# Number of Grade&Zone runners by grade and timespan - 1 xlsx file
# Daily runner flow by week - 1 line graph
# Daily runner flow by weekday - 1 density graph
# Timespan runner flow by workday/weekend - 4 area graphs
# Proportion of runners' preferences to zone - 1 pie chart
# Machines' usage ratio - 1 xlsx file
#
# Author: @JiaXuewei
# Time: 2018.10.17
# Last Update Time: 2018.10.17

thisWeek<-4
dirPath<-"C:/Users/asjxw/Desktop/TimeAnalysis/"
if(!file.exists(dirPath)){ dir.create(dirPath) }
dirPath<-paste("C:/Users/asjxw/Desktop/TimeAnalysis/", "Week", as.character(thisWeek), sep = "", collapse = "")
if(!file.exists(dirPath)){ dir.create(dirPath) }
setwd(dirPath)

library("readxl")
library("xlsx")
library("dplyr")
library("ggplot2")
library("parallel")
library("stringr")
cores<-detectCores()
windowsFonts(myFont = windowsFont("华文细黑"))
filePath<-"C:/Users/asjxw/Desktop/1.xlsx"
sheetsNum<-excel_sheets(filePath)
rawData<-read_xlsx(filePath, col_names = FALSE, sheet = 1)
for(single in sheetsNum[-1]){
  tempData<-read_xlsx(filePath, col_names = FALSE, sheet = single)
  tempData<-tempData[-1,]
  rawData<-rbind(rawData,tempData)
}
rawData<-rawData[-1,]
data<-rawData
colnames(data)<-c("Number","Name","Location","Machine","DateTime")
data$Number<-substr(data$Number, 2, 10)
data[["Date"]]<-substr(data$DateTime, 2, 11)
data[["Time"]]<-substr(data$DateTime, 13, 20)
data[["Weekday"]]<-weekdays(as.Date(data$Date))
data[["Zone"]]<-""
data[["Grade"]]<-substr(data$Number, 1, 2)
data$DateTime<-substr(data$DateTime, 2, 20)
data$Location<-factor(data$Location, levels = c("十号楼", "二号楼", "七号楼", "生物楼", "第二教学楼", "林业楼"))
data$Weekday<-factor(data$Weekday, levels = c("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"))

reverseData<-data[seq(dim(data)[1], 1),]
reverseData<-reverseData[!duplicated(reverseData[c(1, 6)]),]
endTimeData<-reverseData[seq(dim(reverseData)[1], 1),]
endTimeData<-endTimeData[c("Number", "Date", "Time")]
endTimeData$Time<-substr(endTimeData$Time, 1, 5)
colnames(endTimeData)<-c("Number", "Date", "EndTime")

# differences between try["Location"](a list) and try$Location(selected)
# mode(try["Location]) -> "list"
# mode(try$Location) -> "numeric"
returnZone<-function(x) {
  switch(as.character(x["Location"]), 十号楼="西区", 二号楼="西区", 七号楼="西区",
         生物楼="东区", 第二教学楼="东区", 林业楼="东区")
}
cl<-makeCluster(4)
data[["Zone"]]<-parApply(cl, data["Location"], 1, returnZone)
stopCluster(cl)

data[["Week"]]<-floor(difftime(as.Date(data$Date), as.Date("2018-09-17"), units = "weeks"))
data$Week<-as.numeric(as.character(data$Week)) + 1
if (thisWeek != "") { data<-subset(data, Week==thisWeek) }
if (thisWeek == "") { weekAmount<-max(data$Week) } else { weekAmount<-1 }

returnWeek<-function(week) {
  switch(week, "第一周", "第二周", "第三周", "第四周")
}
cl<-makeCluster(4)
data$Week<-parApply(cl, data["Week"], 1, returnWeek)
stopCluster(cl)
data$Week<-factor(data$Week, levels = c("第一周", "第二周", "第三周", "第四周"))

# CleasingData, DataBackup, deduplication
sdata<-data
sdata<-sdata[!duplicated(sdata[c(1, 6)]),]
sdata$Location<-factor(sdata$Location, levels = c("十号楼", "二号楼", "七号楼", "生物楼", "第二教学楼", "林业楼"))
sdata$Time<-substr(sdata$Time, 1, 5)
### set start time time.x and end time time.y
total<-left_join(sdata, endTimeData, by = c("Number", "Date"))


################################ DATA MANIPULATION ####################################


### 每天(星期几), 以东西区分类, 17、18 级细分类的人数统计
monday<-subset(total, total$Weekday == "星期一")
tuesday<-subset(total, total$Weekday == "星期二")
wednesday<-subset(total, total$Weekday == "星期三")
thursday<-subset(total, total$Weekday == "星期四")
friday<-subset(total, total$Weekday == "星期五")
saturday<-subset(total, total$Weekday == "星期六")
sunday<-subset(total, total$Weekday == "星期日")
getTime<-function(weekday, weekdayString, timeString, startTime, endTime) {
  c(weekdayString, timeString,
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "东区" & (weekday$Grade == "17" | weekday$Grade == "18"))),
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "东区" & weekday$Grade == "17")),
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "东区" & weekday$Grade == "18")),
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "西区" & (weekday$Grade == "17" | weekday$Grade == "18"))),
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "西区" & weekday$Grade == "17")),
    length(which(weekday$Time >= startTime & weekday$Time <= endTime & weekday$Zone == "西区" & weekday$Grade == "18")))
}
getAllTime<-function(weekday, weekdayString){
  morning<-getTime(weekday, weekdayString, "06:30-07:30", "06:30", "07:30")
  morning2<-getTime(weekday, weekdayString, "09:00-11:00", "09:00", "11:00")
  afternoon1<-getTime(weekday, weekdayString, "15:30-16:30", "15:30", "16:30")
  afternoon2<-getTime(weekday, weekdayString, "16:30-17:30", "16:30", "17:30")
  afternoon3<-getTime(weekday, weekdayString, "17:30-19:00", "17:30", "19:00")
  return(rbind(morning, morning2, afternoon1, afternoon2, afternoon3))
}
weekdaySummary<-rbind(getAllTime(monday, "Monday"),
      getAllTime(tuesday, "Tuesday"),
      getAllTime(wednesday, "Wednesday"),
      getAllTime(thursday, "Thursday"),
      getAllTime(friday, "Friday"),
      getAllTime(saturday, "Saturday"),
      getAllTime(sunday, "Sunday"))
weekdaySummary<-as.data.frame(weekdaySummary)
colnames(weekdaySummary)<-c("Weekday", "TimeSpan", "East", "East17", "East18", "West", "West17", "West18")
write.xlsx(weekdaySummary, "Summary.xlsx")


workday<-subset(total, total$Weekday != "星期六" & total$Weekday != "星期日")
weekend<-subset(total, total$Weekday == "星期六" | total$Weekday == "星期日")
countTime<-function(data, x1, x2, time) {
  c(time >= data[x1] & time <= data[x2] & data["Grade"] == "17",
    time >= data[x1] & time <= data[x2] & data["Grade"] == "17" & data["Zone"] == "东区",
    time >= data[x1] & time <= data[x2] & data["Grade"] == "17" & data["Zone"] == "西区",
    time >= data[x1] & time <= data[x2] & data["Grade"] == "18", 
    time >= data[x1] & time <= data[x2] & data["Grade"] == "18" & data["Zone"] == "东区",
    time >= data[x1] & time <= data[x2] & data["Grade"] == "18" & data["Zone"] == "西区" )
}

countTime<-function(data, x1, x2, time) {
  c(time >= data[x1] & time <= data[x1] & data["Grade"] == "17",
    time >= data[x1] & time <= data[x1] & data["Grade"] == "17" & data["Zone"] == "东区",
    time >= data[x1] & time <= data[x1] & data["Grade"] == "17" & data["Zone"] == "西区",
    time >= data[x1] & time <= data[x1] & data["Grade"] == "18", 
    time >= data[x1] & time <= data[x1] & data["Grade"] == "18" & data["Zone"] == "东区",
    time >= data[x1] & time <= data[x1] & data["Grade"] == "18" & data["Zone"] == "西区" )
}

setTimeNo<-function(weM, whichDayData, countTime, time, totalDay) {
  result<-apply(whichDayData, 1, countTime, "Time", "EndTime", weM[time])
  cbind(round(length(which(result[1,]) == TRUE)/totalDay), round(length(which(result[2,]) == TRUE)/totalDay),
        round(length(which(result[3,]) == TRUE)/totalDay), round(length(which(result[4,]) == TRUE)/totalDay),
        round(length(which(result[5,]) == TRUE)/totalDay), round(length(which(result[6,]) == TRUE)/totalDay))
}

makeTable<-function(tableName, startTime, mins) {
  tableName<-strptime(startTime, "%H:%M") + 60*0:mins
  tableName<-as.data.frame(tableName)
  colnames(tableName)<-"Var1"
  tableName$Var1<-substr(tableName$Var1, 12, 16)
  return(tableName)
}


### 查看每天的各个时刻跑步人数 ###
# cl<-makeCluster(4)
# workdayTable<-makeTable(mondayTable, "06:30", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, monday, countTime, "Var1", weekAmount*5)))
# colmonday<-cbind(workdayTable, result)
# colnames(colmonday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colmonday, "mondaySummary.xlsx", sheetName = "monday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(tuesdayTable, "06:30", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, tuesday, countTime, "Var1", weekAmount*5)))
# coltuesday<-cbind(workdayTable, result)
# colnames(coltuesday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(coltuesday, "tuesdaySummary.xlsx", sheetName = "tuesday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(saturdayTable, "06:30", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, saturday, countTime, "Var1", weekAmount*5)))
# colsaturday<-cbind(workdayTable, result)
# colnames(colsaturday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colsaturday, "saturdaySummary.xlsx", sheetName = "saturday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(thursdayTable, "06:30", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, thursday, countTime, "Var1", weekAmount*5)))
# colthursday<-cbind(workdayTable, result)
# colnames(colthursday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colthursday, "thursdaySummary.xlsx", sheetName = "thursday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(fridayTable, "06:30", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, friday, countTime, "Var1", weekAmount*5)))
# colfriday<-cbind(workdayTable, result)
# colnames(colfriday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colfriday, "fridaySummary.xlsx", sheetName = "friday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(saturdayTable, "09:00", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, saturday, countTime, "Var1", weekAmount*5)))
# colsaturday<-cbind(workdayTable, result)
# colnames(colsaturday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colsaturday, "saturdaySummary.xlsx", sheetName = "saturday", row.names = FALSE, append = FALSE)
# 
# workdayTable<-makeTable(sundayTable, "09:00", 60*13)
# result<-as.data.frame(t(parApply(cl, workdayTable, 1, setTimeNo, sunday, countTime, "Var1", weekAmount*5)))
# colsunday<-cbind(workdayTable, result)
# colnames(colsunday)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colsunday, "sundaySummary.xlsx", sheetName = "sunday", row.names = FALSE, append = FALSE)
# stopCluster(cl)


#### 横向显示类别
# wdMTable<-makeTable(wdMTable, "06:30", 60)
# result<-as.data.frame(t(parApply(cl, wdMTable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
# colWdM<-cbind(wdMTable, result)
# colnames(colWdM)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colWdM, "TimeSummary.xlsx", sheetName = "WdM", row.names = FALSE, append = FALSE)
# 
# wdATable<-makeTable(wdATable, "15:00", 60*4)
# result<-as.data.frame(t(parApply(cl, wdATable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
# colwdA<-cbind(wdATable, result)
# colnames(colwdA)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colwdA, "TimeSummary.xlsx", sheetName = "wdA", row.names = FALSE, append = FALSE)
# 
# weMTable<-makeTable(weMTable, "09:00", 60*2)
# result<-as.data.frame(t(parApply(cl, weMTable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
# colweM<-cbind(weMTable, result)
# colnames(colweM)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colweM, "TimeSummary.xlsx", sheetName = "weM", row.names = FALSE, append = FALSE)
# 
# weATable<-makeTable(weATable, "15:00", 60*4)
# result<-as.data.frame(t(parApply(cl, weATable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
# colweA<-cbind(weATable, result)
# colnames(colweA)<-c("Var1", "17", "17East", "17West", "18", "18East", "18West")
# write.xlsx(colweA, "TimeSummary.xlsx", sheetName = "weA", row.names = FALSE, append = FALSE)



##### 制作以工作日、周末为分类的 17、18 级分时人流量图表 #####
## 数据汇总
wdMTable<-makeTable(wdMTable, "06:30", 60)
result<-as.data.frame(t(parApply(cl, wdMTable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
wdMTable<-cbind(wdMTable, result[2], "17级")
colnames(wdMTable)<-c("Var1", "Freq", "Grade")
result17<-cbind(wdMTable[1], result[1], "18级")
colnames(result17)<-c("Var1", "Freq", "Grade")
wdMTable<-rbind(wdMTable, result17)

wdATable<-makeTable(wdATable, "15:00", 60*4)
result<-as.data.frame(t(parApply(cl, wdATable, 1, setTimeNo, weekdayData, countTime, "Var1", weekAmount*5)))
wdATable<-cbind(wdATable, result[2], "17级")
colnames(wdATable)<-c("Var1", "Freq", "Grade")
result17<-cbind(wdATable[1], result[1], "18级")
colnames(result17)<-c("Var1", "Freq", "Grade")
wdATable<-rbind(wdATable, result17)

weMTable<-makeTable(weMTable, "09:00", 60*2)
result<-as.data.frame(t(parApply(cl, weMTable, 1, setTimeNo, weekendData, countTime, "Var1", weekAmount*5)))
weMTable<-cbind(weMTable, result[2], "17级")
colnames(weMTable)<-c("Var1", "Freq", "Grade")
result17<-cbind(weMTable[1], result[1], "18级")
colnames(result17)<-c("Var1", "Freq", "Grade")
weMTable<-rbind(weMTable, result17)

weATable<-makeTable(weATable, "15:00", 60*4)
result<-as.data.frame(t(parApply(cl, weATable, 1, setTimeNo, weekendData, countTime, "Var1", weekAmount*5)))
weATable<-cbind(weATable, result[2], "17级")
colnames(weATable)<-c("Var1", "Freq", "Grade")
result17<-cbind(weATable[1], result[1], "18级")
colnames(result17)<-c("Var1", "Freq", "Grade")
weATable<-rbind(weATable, result17)
stopCluster(cl)

## 图表绘制
wdMTable$Var1<-strptime(wdMTable$Var1, "%H:%M")
wdMArea<-ggplot(wdMTable, aes(x = wdMTable$Var1, y = wdMTable$Freq, fill = wdMTable$Grade)) +
  geom_area(alpha = 0.6) +
  scale_x_datetime(date_breaks = "5 min", date_labels = "%H:%M",
                   limits = c(as.POSIXct("06:30", format = "%H:%M"),
                              as.POSIXct("07:30", format = "%H:%M"))) +
  geom_smooth(method = "loess", aes(color = wdMTable$Grade), se = FALSE) +
  labs(title = "工作日 06:30-07:30 时段日均跑步人数", x = "时间", y ="人数", fill = "年级", color = "年级") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.x = element_text(angle = 45, hjust = 1)) +
  scale_y_continuous(limits = c(-1, NA))
ggsave("WdMArea.png", wdMArea, width = 10, height = 6)

wdATable$Var1<-strptime(wdATable$Var1, "%H:%M")
wdAArea<-ggplot(wdATable, aes(x = wdATable$Var1, y = wdATable$Freq, fill = wdATable$Grade)) +
  geom_area(alpha = 0.6) +
  scale_x_datetime(date_breaks = "10 min", date_labels = "%H:%M",
                   limits = c(as.POSIXct("15:30", format = "%H:%M"),
                              as.POSIXct("19:00", format = "%H:%M"))) +
  geom_smooth(method = "loess", aes(color = wdATable$Grade), se = FALSE) +
  labs(title = "工作日 15:30-19:00 时段日均跑步人数", x = "时间", y ="人数", fill = "年级", color = "年级") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.x = element_text(angle = 45, hjust = 1)) +
  scale_y_continuous(limits = c(-1, NA))
ggsave("WdAArea.png", wdAArea, width = 10, height = 6)

weMTable$Var1<-strptime(weMTable$Var1, "%H:%M")
weMArea<-ggplot(weMTable, aes(x = weMTable$Var1, y = weMTable$Freq, fill = weMTable$Grade)) +
  geom_area(alpha = 0.6) +
  scale_x_datetime(date_breaks = "5 min", date_labels = "%H:%M",
                   limits = c(as.POSIXct("09:00", format = "%H:%M"),
                              as.POSIXct("11:00", format = "%H:%M"))) +
  geom_smooth(method = "loess", aes(color = weMTable$Grade), se = FALSE) +
  labs(title = "周末 09:00-11:00 时段日均跑步人数", x = "时间", y ="人数", fill = "年级", color = "年级") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.x = element_text(angle = 45, hjust = 1)) +
  scale_y_continuous(limits = c(0, NA))
ggsave("WeMArea.png", weMArea, width = 10, height = 6)

weATable$Var1<-strptime(weATable$Var1, "%H:%M")
weAArea<-ggplot(weATable, aes(x = weATable$Var1, y = weATable$Freq, fill = weATable$Grade)) +
  geom_area(alpha = 0.6) +
  scale_x_datetime(date_breaks = "10 min", date_labels = "%H:%M",
                   limits = c(as.POSIXct("15:30", format = "%H:%M"),
                              as.POSIXct("19:00", format = "%H:%M"))) +
  geom_smooth(method = "loess", aes(color = weATable$Grade), se = FALSE) +
  labs(title = "周末 15:30-19:00 时段日均跑步人数", x = "时间", y ="人数", fill = "年级", color = "年级") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.x = element_text(angle = 45, hjust = 1)) +
  scale_y_continuous(limits = c(-1, NA))
ggsave("WeAArea.png", weAArea, width = 10, height = 6)


################################ INDEPENDENT GRAPHS ####################################


##### 制作刷卡机负载对比图 #####
machinePlot<-ggplot(data, aes(Machine, fill = Location)) +
  geom_bar() +
  coord_polar(theta = "x", direction = 1) +
  labs(title = "刷卡机负载", x = "刷卡机", fill = "刷卡点") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.y = element_blank(),
        axis.ticks = element_blank(),
        axis.title.y = element_blank())
ggsave("MachinePlot.png", machinePlot, width = 8, height = 6)


##### 输出刷卡机刷卡量 Excel 表格 #####
machineTable<-as.data.frame(table(data$Machine, data$Location))
machineTable$Var2<-factor(machineTable$Var2, levels = c("十号楼", "二号楼", "七号楼", "生物楼", "第二教学楼", "林业楼"), ordered = TRUE)
machineTable<-machineTable[!machineTable[3] == 0,]
colnames(machineTable)<-c("刷卡机", "刷卡点", "刷卡次数")
machineTable<-machineTable[order(machineTable$刷卡点, -machineTable$刷卡次数),]
westMachineTable<-machineTable[1:6,]
eastMachineTable<-machineTable[7:length(machineTable$刷卡机),]
write.xlsx(westMachineTable, "MachineTable.xlsx", row.names = FALSE, sheetName = "West", append = FALSE)
write.xlsx(eastMachineTable, "MachineTable.xlsx", row.names = FALSE, sheetName = "East", append = TRUE)


##### 按 17、18 级分类的东西区跑步人数比例图 #####
myLabel<-c("东区:17级", "东区:18级", "西区:17级", "西区:18级")
validStudentNumber<-length(which(sdata$Grade!=17 | sdata$Grade!=18))
zoneCount<-data.frame(myLabel, c(round(length(which(sdata$Zone == "东区" & sdata$Grade == 17 ))/validStudentNumber * 100, 2),
                                 round(length(which(sdata$Zone == "东区" & sdata$Grade == 18 ))/validStudentNumber * 100, 2),
                                 round(length(which(sdata$Zone == "西区" & sdata$Grade == 17 ))/validStudentNumber * 100, 2),
                                 round(length(which(sdata$Zone == "西区" & sdata$Grade == 18 ))/validStudentNumber * 100, 2)))
colnames(zoneCount)<-c("Var1", "Freq")
myLabel<-paste(myLabel, "(", c(round(length(which(sdata$Zone == "东区" & sdata$Grade == 17 ))/validStudentNumber * 100, 2),
                               round(length(which(sdata$Zone == "东区" & sdata$Grade == 18 ))/validStudentNumber * 100, 2),
                               round(length(which(sdata$Zone == "西区" & sdata$Grade == 17 ))/validStudentNumber * 100, 2),
                               round(length(which(sdata$Zone == "西区" & sdata$Grade == 18 ))/validStudentNumber * 100, 2)), "%) ", sep = "")
sdata$Grade<-factor(sdata$Grade, levels = c(17, 18))
zonePie<-ggplot(zoneCount, aes(x = "", y = Freq, fill = Var1)) +
  geom_bar(stat = "identity", width = 1) +
  coord_polar(theta = "y") +
  labs(x = "", y = "", title = "东西区跑步人数比例") +
  theme(axis.ticks = element_blank(),
        text = element_text(family = "myFont", size = 18, vjust = 5),
        legend.title = element_blank(),
        # legend.position = "bottom",
        axis.text.x = element_blank()) +
  scale_fill_discrete(breaks = c("东区:17级", "东区:18级", "西区:17级", "西区:18级"),labels = myLabel)

ggsave("ZonePie.png", zonePie, width = 8, height = 6)


##### 周平均跑步人数密度图 #####
everyday<-table(sdata$Weekday)
everyday<-as.data.frame(everyday)
    # geom_line/area 等要求寻找组, 则 Group 必须 = 1, 方可将目标设为同一组
densityArea<-ggplot(everyday, aes(x = Var1, y = Freq/sum(everyday$Freq), group = 1)) +
  geom_area(alpha = 0.6, fill="#51A8DD") +
  labs(title = "周平均跑步人数密度", x = "星期", y = "密度") +
  theme(text = element_text(family = "myFont", size = 18),
        axis.text.x = element_text(angle = 30, hjust = 1))
ggsave("DensityArea.png", densityArea, width = 8, height = 5)


##### 分周次跑步人数条形图 #####
week<-table(sdata$Weekday, sdata$Week)
week<-as.data.frame(week)
weekLine<-ggplot(week, aes(x = Var1, y = Freq, colour = Var2, group = Var2)) +
  geom_line(size = 1) +
  labs(x = "星期", y = "总人数", title = "分周次跑步人数", colour = "周次") +
  theme(text = element_text(family = "myFont", size = 18))
ggsave("WeekLine.png", weekLine, height = 5, width = 8)
