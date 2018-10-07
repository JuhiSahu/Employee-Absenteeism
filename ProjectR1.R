rm(list = ls())
setwd("E:/Study/edWisor Stuff/Project1")

#load the data
install.packages("xlsx")
library(xlsx)
data1 = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1, header = TRUE)
ds = data1
#data1=ds #copy
str(data1)

#PREPROCESSING
sapply(data1, function(x) sum(is.na(x)))

#converting data types
data1$Month.of.absence = factor(data1$Month.of.absence)
data1$Day.of.the.week= factor(data1$Day.of.the.week)
data1$Seasons = factor(data1$Seasons)
data1$Disciplinary.failure = factor(data1$Disciplinary.failure)
data1$Education = factor(data1$Education)
data1$Social.drinker = factor(data1$Social.drinker)
data1$Social.smoker = factor(data1$Social.smoker)

#col2
data1$Reason.for.absence[data1$Reason.for.absence == 0] = '26'
sum(is.na(data1$Reason.for.absence))
data1$Reason.for.absence = factor(data1$Reason.for.absence)

#col18
data1$Weight[is.na(data1$Weight)] = mean(data1$Weight, na.rm = T)
data1$Weight = as.integer(data1$Weight)

#col19
data1$Height[is.na(data1$Height)] = mean(data1$Height, na.rm = T)
data1$Height = as.integer(data1$Height)

#col20
data1$Body.mass.index[is.na(data1$Body.mass.index)] = (data1$Weight *10000) /(data1$Height**2)
data1$Body.mass.index = as.integer(data1$Body.mass.index)

library(DMwR)
data1 = knnImputation(data1, k=3)


#------------------------------------------VISVUALIZATION-----------------------------------
library(ggplot2)
library(scales)
library(psych)

#?ggplot()
#ReasonsCount
ggplot(data1 , aes_string(x=data1$Reason.for.absence)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("Reason") + ylab("Count") + ggtitle("Reason distribution") + theme(text = element_text(size = 15))
#non ICD reasons have higher count

#PetCount
ggplot(data1 , aes_string(x=data1$Pet)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("No Of Pets") + ylab("Count") + ggtitle("Pet distribution")

plot(Absenteeism.time.in.hours ~ Pet , data = data1)
#people with atleast one pet show less absentism in hours

#transportation expenses
ggplot(data1 , aes_string(x=data1$Transportation.expense)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("Transportation expense") + ylab("Count") + ggtitle("Transportation expanse distribution")

plot(Absenteeism.time.in.hours ~ Transportation.expense, data = data1)

#Drinker
ggplot(data1 , aes_string(x=data1$Social.drinker)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("Drinker") + ylab("Count") + ggtitle("Drinker distribution")

plot(Absenteeism.time.in.hours ~ Social.drinker , data = data1)
#People who are social drinkers tend to be more absent

#Season
ggplot(data1 , aes_string(x=data1$Seasons)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("Seasons") + ylab("Count") + ggtitle("Season distribution")

#Month of absence
ggplot(data1 , aes_string(x= data1$Month.of.absence)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("Month") + ylab("Count") + ggtitle("Month distribution")
#month 3rd have higher absentism

#dayOfWeek
ggplot(data1 , aes_string(x=data1$Day.of.the.week)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("DayOfWeek") + ylab("Count") + ggtitle("Day distribution")

plot(Absenteeism.time.in.hours ~ Day.of.the.week , data = data1)
#people tend to be least absent on thursdays.

#disciplinary failure
ggplot(data1 , aes_string(x=data1$Disciplinary.failure)) + geom_bar(stat = "count", fill = "DarkslateBlue") + 
  xlab("DayOfWeek") + ylab("Count") + ggtitle("Season distribution")

plot(Absenteeism.time.in.hours ~ Disciplinary.failure , data = data1)
#higher number of people do not have disciplinary failure


#---------------------------------------OUTLIER ANALYSIS------------------------------------

#Boxplotting
numeric_index = sapply(data1, is.numeric)
numeric_data = data1[,numeric_index]
cnames = colnames(numeric_data)

for (i in 1:length(cnames)){
  assign(paste0("gn", i), ggplot(aes_string(y = cnames[i], x = "Absenteeism.time.in.hours"), data = subset(data1)) +
           stat_boxplot(geom = "errorbar", width = 0.5) + 
           geom_boxplot(outlier.colour = "red", fill = "grey", outlier.shape = 18, outlier.size = 1, notch = FALSE) + 
           theme(legend.position = "bottom") + labs(y = cnames[i], x="abs hours") + ggtitle(paste("Boxplot of abs hours for", cnames[i])))
}


#plotting plots together
gridExtra::grid.arrange(gn2, gn3, ncol = 2)
gridExtra::grid.arrange(gn4, gn5, ncol = 2)
gridExtra::grid.arrange(gn6, gn7, ncol = 2)
gridExtra::grid.arrange(gn8, gn9, ncol = 2)
gridExtra::grid.arrange(gn10, gn11, ncol = 2)
gridExtra::grid.arrange(gn12, gn13, ncol = 2)


#replace outliers with NA and impute
for(i in cnames) {
   print(i)
  val = data1[,i][data1[,i] %in% boxplot.stats(data1[,i]) $out]
  
  print(length(val))
  data1[,i][data1[,i] %in% val] = NA
}

data1 = knnImputation(data1, k=3)
#sapply(data1, function(x) sum(is.na(x)))


#-------------------------------------FEATURE SELECTION--------------------------------

#correlation plot
library(corrgram)

round(cor(numeric_data),2)

corrgram(data1[, numeric_index], order = F, upper.panel = panel.pie, text.panel = panel.txt, main = "correlation plot")

data1_new = subset(data1, select=-c(Height, Weight, Distance.from.Residence.to.Work)) 
#removing weight, Distance.from.Residence and height

#Anova test

#season of absence
AnovaModel_season =(lm(Absenteeism.time.in.hours ~ Seasons, data = data1))
summary(AnovaModel_season) #remove

#Reason of absence
AnovaModel_reason=(lm(Absenteeism.time.in.hours ~ Reason.for.absence, data = data1))
summary(AnovaModel_reason) #keep

#month of week
AnovaModel_month=(lm(Absenteeism.time.in.hours ~ Month.of.absence, data = data1))
summary(AnovaModel_month) #keep

#day
AnovaModel_day=(lm(Absenteeism.time.in.hours ~ Day.of.the.week, data = data1))
summary(AnovaModel_day) #remove

#Disciplinary failure
AnovaModel_disciplinary=(lm(Absenteeism.time.in.hours ~ Disciplinary.failure, data = data1))
summary(AnovaModel_disciplinary) #remove

#Education
AnovaModel_education=(lm(Absenteeism.time.in.hours ~ Education, data = data1))
summary(AnovaModel_education) #remove

#drinker
AnovaModel_drinker=(lm(Absenteeism.time.in.hours ~ Social.drinker, data = data1))
summary(AnovaModel_drinker) #keep

#smoker
AnovaModel_smoker=(lm(Absenteeism.time.in.hours ~ Social.smoker, data = data1))
summary(AnovaModel_smoker) #keep

data1_new = subset(data1_new, select=-c(ID, Seasons, Day.of.the.week, Education, Disciplinary.failure))


#-------------FEATURE SCALING---------------------

#normality check
hist(data1_new$Transportation.expense)
hist(data1_new$Service.time)
hist(data1_new$Age)
hist(data1_new$Work.load.Average.day.)
hist(data1_new$Body.mass.index)
hist(data1_new$Absenteeism.time.in.hours)
#since data is not normally distributed of any column we will use normalization

cnames = cnames[-c(1,3,10,11)]

for (i in cnames){
  print(i)
  data1_new[,i] = (data1_new[,i]-min(data1_new[,i]))/ (max(data1_new[,i])-min(data1_new[,i]))
}

#-------------MODELING----------------------------

#sampling
train_index = sample(1:nrow(data1_new), 0.8*nrow(data1_new))
data1_train = data1_new[train_index,] 
data1_test = data1_new[-train_index,]

#LINEAR REGRESSION
library(rpart)
library(MASS)

#check multicollinearity
install.packages("usdm")
library(usdm)

vif(numeric_data[,-13])
vifcor(numeric_data[,-13], th = 0.9) 
#no variable from the 12 input variables has collinearity problem.

#creating dummies
install.packages("dummies")
library(dummies)

#?dummy.data.frame()
df_new = dummy.data.frame(data1_new, sep = '.')
dim(df_new)
colnames(df_new)

#sampling 
train_index = sample(1:nrow(df_new), 0.8*nrow(df_new))
df_train = df_new[train_index,] 
df_test = df_new[-train_index,]

#run regression model
lm_model11 = lm(Absenteeism.time.in.hours~. , data = df_train)
summary(lm_model11)
#R square= 0.47
#Adjusted R square = 0.42

#predict
predictions_LR = predict(lm_model11, df_test[,-69])

#Calculate MAPE

#mape = function(y, yhat){
 # mean(abs((y-yhat)/y))*100
#}

#mape(df_test[,69], predictions_LR)

#library(DMwR)
regr.eval(df_test[-69], predictions_LR, stats = c('mse','rmse','mape','mae'))
#rmse = 30.35




#Decision Treet
install.packages("rpart.plot")
library(rpart.plot)
library(rpart)
fit = rpart(Absenteeism.time.in.hours~. , data = data1_train, method = "anova")
plt = rpart.plot(fit, type = 3, digits = 2, fallen.leaves = TRUE)

Predict_DT = predict(fit, data1_test[,-20])

#accuracy
regr.eval(data1_test[,13], Predict_DT, stats = c('mae', 'mse', 'rmse', 'mape'))

#mae= 1.90
#mse= 8.08
#rmse= 2.84
#mape= Inf
#accuracy = 97.16


#Random Forest
install.packages("randomForest")
library(randomForest)
library(inTrees)
RF_model = randomForest(Absenteeism.time.in.hours~. , data1_train, importance = TRUE, ntree = 500)
treeList = RF2List(RF_model)
#error plotting
plot(RF_model)

#extract rules
exec = extractRules(treeList, data1_train[,-13])
#visvualise rules
exec[1:2,]

#make rules more readable
readableRules = presentRules(exec, colnames(data1_train))

#Rule matrix
ruleMatrix = getRuleMetric(exec, data1_train[,-13], data1_train$Absenteeism.time.in.hours)

#predict test data using RF model
RF_predict = predict(RF_model, data1_test[,-13])

#evaluate performance
postResample(RF_predict, data1_test$Absenteeism.time.in.hours)
#rmse = 2.87
#rsquare = 0.25
#mae = 2.05 
#accuracy = 97.13


#-----------------------------MONTHLY LOSS

new = subset(data1, select = c(Month.of.absence, Service.time, Absenteeism.time.in.hours, Work.load.Average.day. ))

#Work loss = ((Work load per day/ service time)* Absenteeism hours)

new["loss"]=with(new,((new[,4]*new[,3])/new[,2]))

for(i in 1:12)
{
  d1=new[which(new["Month.of.absence"]==i),]
  cat("\n month:",i, sum(d1$loss))
  
}





