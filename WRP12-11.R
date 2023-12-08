# load packages
library(tidyverse)
library(lubridate)
library(scales)
library(extrafont)
library(openxlsx)


# set strings as factors to false
options(stringsAsFactors = FALSE)

###############################################################################
#INPUT
###############################################################################

#catchment data
catch_name <- "EmpGentingKlang"
catch_A <- 76 #km2

#rainfall data
station_no <- "3217002" #station number
station_name <- "Emp Genting Klang"
#rainfall data end before year
end_year <- "2022"

#evaporation station
ev_stn <- "JPS Ampang"

#find Q for percentage n
Qn <- 95

#baseflow input (from neighbouring catchment)
bf_stn <- "3217401" #Sg Gombak di Damsite
bf_Q_min <- 0.08 #m3/s (1983-1991)
bf_catch_A <- 84.7 #km2


#parameters
water_holding <- 250 #mm
initial_soil_moisture <- 250 #mm
initial_avail_water_runoff <- 200 #mm
soil_mois_retention_a <- 249.5 #WRP 6, 2.3.3
soil_mois_retention_b <- -0.004 #WRP 6, 2.3.3
recession_K <- 0.98
#surface_runoff_rate <- 0.1 #HSG = B

#minimum count per year for analysis
min_cnt <- 365


#calculation
bf_specific_Q <- bf_Q_min/bf_catch_A #m3/s/km2
bf_Q_cumecs <- bf_specific_Q*catch_A #m3/s
bf_Q_mm_day <- bf_Q_cumecs/(catch_A*10^6)*1000*(24*60*60) #mm/day
Qb <- bf_Q_mm_day #mm/day
  


###########################
#INPUT FILES
###########################

#set working directory
#set filename
filename2 <- paste0("WRP12_", station_no, "_", catch_name)
# get current working directory
working_dir <- getwd()
dir.create(filename2)
setwd(filename2)


###########################
##OPTIONS (rainfall station data or from database) #############

#input format in c(Date, Depth) in csv (case sensitive), station_no is station number, Depth in mm
#station_no <- strsplit("3014091.csv", "\\.")[[1]][1]
raindata <- read.csv(file=paste0(working_dir,"/", station_no, ".csv"),
                     header = TRUE, sep=",") #depends on working dir

#rename columns, format date
colnames(raindata) <- c("Date", "Depth")

#check field name
head(raindata, 3) 

raindata$Date <- as.Date(raindata$Date, format = "%d/%m/%Y")

#############
#from rainfall database
rain_db <- read.csv(file="F:/Documents/2019/20160506 Tangki NAHRIM/Data/SQL/rainfall_pm_clean6d.csv",
                    header = TRUE, sep=",")
#check field name
head(rain_db, 3)

#extract station data
raindata <- rain_db %>%
  filter(stn_no == station_no) %>% #check field name
  select(date, depth) %>%
  arrange(date)

#rename columns, format date
colnames(raindata) <- c("Date", "Depth")

#raindata$Date <- as.Date(raindata$Date, format = "%d/%m/%Y")
raindata$Date <- as.Date(raindata$Date, format = "%Y-%m-%d")




###########################
#EVAPORATION (WRP 5)
ev_data <- read.csv(file=paste0(working_dir,"/", ev_stn, ".csv"),
                    header = TRUE, sep=",") #depends on working dir
#check field name
head(ev_data, 3)
#add column for days in month
ev_data <- ev_data %>%
  mutate(days_mth = days_in_month(Month),
         PE = Eto/days_mth)
#rename columns, format date
#ev_data$Date <- as.Date(ev_data$Date, format = "%d/%m/%Y")
#same as spreadsheet, Feb has 29 days
ev_data[2, 3] <- 29
ev_data <- ev_data %>%
  mutate(PE = Eto/days_mth)


###############################################################################
#CALCULATION
###############################################################################

###########################
##PRECIPITATION
###########################


# AGGREGATE DATA
# add a year column to data.frame
raindata <- raindata %>%
  mutate(Year = year(Date))
# add a month column to data.frame
raindata <- raindata %>%
  mutate(Month = month(Date))


# MISSING DATA
## remove years with missing data
## count non-missing data
raindata_cnt <- raindata %>%
  group_by(Year) %>%
  summarise(sum_precip = sum(Depth, na.rm = T), cnt = sum(!is.na(Depth)))

## select years without missing data
raindata_cnt_list <- raindata_cnt %>%
  filter(cnt >= min_cnt)

## select data
raindata_sel <- raindata %>%
  right_join(raindata_cnt_list, by = "Year") %>%
  select(Date, Year, Month, Depth)

## replace NA in rainfall data with 0
raindata_sel[is.na(raindata_sel)] <- 0



# EVAPORATION DATA
# add evaporation to data.frame
sim_data <- left_join(raindata_sel, ev_data, by = "Month")
  

# CALCULATION
# calculate the sum precipitation for each year
precip_sum_yr <- raindata_sel %>%
  filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  group_by(Year) %>%
  summarise(sum_precip = sum(Depth, na.rm = T))

#calculate surface runoff rate, fs from avg rainfall
avg_yr_precip <- mean(precip_sum_yr$sum_precip)
surface_runoff_rate <- ifelse(avg_yr_precip < 2000, 
                              0.01,
                              ifelse(avg_yr_precip < 2500,
                                     0.05,
                                     0.1))

# calculate the sum precipitation for each month, each year
precip_sum_mth <- raindata_sel %>%
  #filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  group_by(Month, Year) %>%
  summarise(sum_precip = sum(Depth)) %>%
  group_by(Month) %>%
  summarise(mth_precip = mean(sum_precip, na.rm = T))

#find max and min year
max_year <- max(precip_sum_yr$Year)
min_year <- min(precip_sum_yr$Year)
total_year <- max_year - min_year + 1
year_w_data <- length(unique(precip_sum_yr$Year))
day_w_data <- length(unique(raindata_sel$Date))
angle_precip <- ifelse(total_year > 20, 90, 0)

###########################
##WRP12
###########################
#initialization for dataframe calculation
#P - Adjusted For Runoff Rate, fs (mm)
sim_data$fs <- 0
#Soil Moisture (mm)
sim_data$SM <- 0
#Water Surplus
sim_data$W_S <- 0 #NA?
#Water Deficit
sim_data$W_D <- 0
#Calculated Runoff ,CRO (mm)
sim_data$CRO <- 0
#Available Water for Runoff ,AWR (mm)
sim_data$AWR <- 0
sim_data$MS_i_1 <- 0
sim_data$AWR_i_1 <- 0
sim_data$P <- 0
sim_data$MS_i <- 0
sim_data$dMS_i <- 0
sim_data$AE <- 0
sim_data$WD <- 0
sim_data$WS <- 0
sim_data$AWR_i <- 0
sim_data$SWR <- 0
sim_data$BF <- 0
sim_data$CRO_mm <- 0
sim_data$CRO_cumecs <- 0


###########################
#spreadsheet calculation 


for (i in 1:nrow(sim_data)) {
  if(i==1) {
    
    sim_data$fs[i] = sim_data$Depth[i]*(1 - surface_runoff_rate)
    sim_data$SM[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                            ifelse(initial_soil_moisture+(sim_data$fs[i]-sim_data$PE[i]) > water_holding,
                                                          water_holding,
                                                          initial_soil_moisture + (sim_data$fs[i]-sim_data$PE[i])),
                            soil_mois_retention_a*exp(soil_mois_retention_b*(
                              (log(initial_soil_moisture/soil_mois_retention_a)/
                                 soil_mois_retention_b)+abs(sim_data$PE[i]-sim_data$fs[i]))))
    sim_data$W_S[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                             ifelse(initial_soil_moisture + (sim_data$fs[i] - sim_data$PE[i]) < water_holding,
                                    0,
                                    water_holding + (sim_data$fs[i] - sim_data$PE[i]) - water_holding),
                             NA)
    sim_data$W_D[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                             NA,
                             -(sim_data$PE[i])+
                               (sim_data$fs[i]+
                                  (initial_soil_moisture - soil_mois_retention_a*
                                     exp(soil_mois_retention_b*
                                           ((log(initial_soil_moisture/soil_mois_retention_a)/
                                               soil_mois_retention_b)+
                                              abs(sim_data$PE[i]-sim_data$fs[i])
                                            )))))
    sim_data$CRO[i] = ifelse((1-recession_K)*initial_avail_water_runoff+surface_runoff_rate*sim_data$Depth[i]
                             <= Qb,
                             max(Qb, 
                                 surface_runoff_rate*sim_data$Depth[i], na.rm = T),
                             max((1-recession_K)*initial_avail_water_runoff + surface_runoff_rate*sim_data$Depth[i],
                                 surface_runoff_rate*sim_data$Depth[i]))
    sim_data$AWR[i] = ifelse(is.na(sim_data$W_S[i]),
                             ifelse(0+(1-recession_K)*
                                      initial_avail_water_runoff - Qb +
                                      sim_data$Depth[i]*surface_runoff_rate
                                    < 0,
                                    initial_avail_water_runoff - Qb + sim_data$Depth[i]*surface_runoff_rate,
                                    0+initial_avail_water_runoff - sim_data$CRO[i] + sim_data$Depth[i]*surface_runoff_rate),
                             ifelse(sim_data$W_S[i] + 
                                      (1 - recession_K)*initial_avail_water_runoff - Qb +
                                      sim_data$Depth[i]*surface_runoff_rate <
                                      0,
                                    sim_data$W_S[i] + initial_avail_water_runoff - 
                                      Qb + sim_data$Depth[i]*surface_runoff_rate,
                                    sim_data$W_S[i]+initial_avail_water_runoff - sim_data$CRO[i] +
                                      sim_data$Depth[i]*surface_runoff_rate))
    sim_data$MS_i_1[i] = initial_soil_moisture
    sim_data$AWR_i_1[i] = initial_avail_water_runoff
    sim_data$P[i] = sim_data$fs[i]
    sim_data$MS_i[i] = sim_data$SM[i]
    sim_data$dMS_i[i] = sim_data$MS_i[i] - sim_data$MS_i_1[i]
    sim_data$AE[i] = ifelse(sim_data$P[i]>sim_data$PE[i],
                            sim_data$PE[i],
                            sim_data$P[i] + abs(sim_data$dMS_i[i]))
    sim_data$WD[i] = sim_data$PE[i] - sim_data$AE[i]
    sim_data$WS[i] = ifelse(is.na(sim_data$W_S[i]), 0, sim_data$W_S[i])
    
    sim_data$SWR[i] = sim_data$Depth[i]*surface_runoff_rate
    sim_data$BF[i] = Qb
    sim_data$CRO_mm[i] = ifelse((1-recession_K)*sim_data$AWR_i_1[i]+sim_data$SWR[i] <= sim_data$BF[i],
                                max(sim_data$BF[i], 
                                    sim_data$SWR[i], na.rm = T),
                                max((1-recession_K)*sim_data$AWR_i_1[i]+sim_data$SWR[i], 
                                    sim_data$SWR[i], na.rm = T))
    sim_data$AWR_i[i] = sim_data$AWR_i_1[i] + sim_data$WS[i] - sim_data$CRO_mm[i] + sim_data$SWR[i]
    sim_data$CRO_cumecs[i] = sim_data$CRO_mm[i]/1000*(catch_A*10^6)/(24*60*60)
    
  } else {
    

    sim_data$fs[i] = sim_data$Depth[i]*(1 - surface_runoff_rate)
    sim_data$SM[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                            ifelse(sim_data$SM[i-1]+(sim_data$fs[i] - sim_data$PE[i]) 
                                   > water_holding,
                                   water_holding,
                                   sim_data$SM[i-1]+(sim_data$fs[i] - sim_data$PE[i])),
                            soil_mois_retention_a*exp(soil_mois_retention_b*(
                              (log(sim_data$SM[i-1]/soil_mois_retention_a)/
                                 soil_mois_retention_b)+abs(sim_data$PE[i]-sim_data$fs[i]))))
    sim_data$W_S[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                             ifelse(sim_data$SM[i-1] + (sim_data$fs[i] - sim_data$PE[i]) < water_holding,
                                    0,
                                    sim_data$SM[i-1] + (sim_data$fs[i] - sim_data$PE[i]) - water_holding),
                             NA)
    sim_data$W_D[i] = ifelse(sim_data$fs[i] > sim_data$PE[i],
                             NA,
                             -(sim_data$PE[i])+
                               (sim_data$fs[i]+
                                  (sim_data$SM[i-1] - soil_mois_retention_a*
                                     exp(soil_mois_retention_b*
                                           ((log(sim_data$SM[i-1]/soil_mois_retention_a)/
                                               soil_mois_retention_b)+
                                              abs(sim_data$PE[i]-sim_data$fs[i])
                                           )))))
    sim_data$CRO[i] = ifelse((1-recession_K)*sim_data$AWR[i-1]+surface_runoff_rate*sim_data$Depth[i]
                             <= Qb,
                             max(Qb, 
                                 surface_runoff_rate*sim_data$Depth[i], na.rm = T),
                             max((1-recession_K)*sim_data$AWR[i-1] + surface_runoff_rate*sim_data$Depth[i],
                                 surface_runoff_rate*sim_data$Depth[i]))
    sim_data$AWR[i] = ifelse(is.na(sim_data$W_S[i]),
                             ifelse(0+(1-recession_K)*
                                      max(0, sim_data$AWR[i-1], na.rm = T) 
                                    - Qb + sim_data$Depth[i]*surface_runoff_rate
                                    < 0,
                                    sim_data$AWR[i-1] - Qb + sim_data$Depth[i]*surface_runoff_rate,
                                    0+sim_data$AWR[i-1] - sim_data$CRO[i] + sim_data$Depth[i]*surface_runoff_rate),
                             ifelse(sim_data$W_S[i] + 
                                      (1 - recession_K)*max(0, sim_data$AWR[i-1], na.rm = T) - Qb +
                                      sim_data$Depth[i]*surface_runoff_rate <
                                      0,
                                    sim_data$W_S[i] + sim_data$AWR[i-1] - 
                                      Qb + sim_data$Depth[i]*surface_runoff_rate,
                                    sim_data$W_S[i]+sim_data$AWR[i-1] - sim_data$CRO[i] +
                                      sim_data$Depth[i]*surface_runoff_rate))
    sim_data$MS_i_1[i] = sim_data$MS_i[i-1]
    sim_data$AWR_i_1[i] = sim_data$AWR_i[i-1]
    sim_data$P[i] = sim_data$fs[i]
    sim_data$MS_i[i] = sim_data$SM[i]
    sim_data$dMS_i[i] = sim_data$MS_i[i] - sim_data$MS_i_1[i]
    sim_data$AE[i] = ifelse(sim_data$P[i]>sim_data$PE[i],
                            sim_data$PE[i],
                            sim_data$P[i] + abs(sim_data$dMS_i[i]))
    sim_data$WD[i] = sim_data$PE[i] - sim_data$AE[i]
    sim_data$WS[i] = ifelse(is.na(sim_data$W_S[i]), 0, sim_data$W_S[i])
    
    sim_data$SWR[i] = sim_data$Depth[i]*surface_runoff_rate
    sim_data$BF[i] = Qb
    sim_data$CRO_mm[i] = ifelse((1-recession_K)*sim_data$AWR_i_1[i]+sim_data$SWR[i] <= sim_data$BF[i],
                                max(sim_data$BF[i], 
                                    sim_data$SWR[i], na.rm = T),
                                max((1-recession_K)*sim_data$AWR_i_1[i]+sim_data$SWR[i], 
                                    sim_data$SWR[i], na.rm = T))
    sim_data$AWR_i[i] = sim_data$AWR_i_1[i] + sim_data$WS[i] - sim_data$CRO_mm[i] + sim_data$SWR[i]
    sim_data$CRO_cumecs[i] = sim_data$CRO_mm[i]/1000*(catch_A*10^6)/(24*60*60)
    
  }

}

###########################
#summarise data

# calculate the sum CRO (mm) for each year
CRO_mm_sum_yr <- sim_data %>%
  filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  select(Year, CRO_mm) %>%
  group_by(Year) %>%
  summarise(sum_CRO_mm = sum(CRO_mm)) 

# calculate the sum CRO (mm) for each month, each year
CRO_mm_sum_mth <- sim_data %>%
  filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  select(Year, Month, CRO_mm) %>%
  group_by(Year, Month) %>%
  summarise(sum_CRO_mm = sum(CRO_mm)) %>%
  group_by(Month) %>%
  summarise(mth_CRO_mm = mean(sum_CRO_mm))


# calculate the sum CRO (m3/s) for each year
CRO_cumecs_sum_yr <- sim_data %>%
  filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  select(Year, CRO_cumecs) %>%
  group_by(Year) %>%
  summarise(sum_CRO_cumecs = sum(CRO_cumecs)) 

# calculate the sum CRO (m3/s) for each month, each year
CRO_cumecs_sum_mth <- sim_data %>%
  filter(Date < as.Date(paste0(end_year, "-01-01"))) %>% #remove for different stations
  select(Year, Month, CRO_cumecs) %>%
  group_by(Year, Month) %>%
  summarise(sum_CRO_cumecs = sum(CRO_cumecs)) %>%
  group_by(Month) %>%
  summarise(mth_CRO_cumecs = mean(sum_CRO_cumecs))


###############################################################################
#FDC
###############################################################################

#number of records
prob_nrow <- nrow(sim_data)

#create FDC data table
FDC_data <- sim_data %>%
  select(Year, Month, CRO_cumecs) %>%
  arrange(desc(CRO_cumecs)) %>%
  mutate(rank = dense_rank(desc(CRO_cumecs))) %>%
  mutate(prob = rank/(prob_nrow+1)*100) %>%
  mutate(id = row_number())

#create FDC table
FDC_table <- data.frame("Percentage" = c(1:100))
FDC_table <- FDC_table %>%
  mutate(Tens = as.integer(Percentage/10)*10,
         Ones = Percentage %% 10)

#find matching values
j <- 1
FDC_table$index <- 0
for (j in 1:nrow(FDC_table)) {
  FDC_table$index[j] <- which.min(abs(FDC_data$prob - FDC_table$Percentage[j]))
}

#match index of matching values
FDC_table <- FDC_table %>%
  left_join(FDC_data, by = c("index" = "id"))

#pivot
FDC_table_pv <- FDC_table %>%
  select(Tens, Ones, CRO_cumecs) %>%
  spread(Ones, CRO_cumecs)


#in MLD
FDC_table_MLD <- FDC_table %>%
  select(Tens, Ones, CRO_cumecs) %>%
  mutate(CRO_MLD = CRO_cumecs*1000/10^6*24*60*60) %>%
  select(Tens, Ones, CRO_MLD) %>%
  spread(Ones, CRO_MLD)

#extract Qn
Qn_ex <- FDC_table[Qn, "CRO_cumecs"] #column CRO_cumecs

###########################
#FDC monthly

#loop

#create df to store Qn_ex
Qn_ex_mth_df <- NULL
list_FDC_mth_cumecs <- list()
list_FDC_mth_MLD <- list()

for (a in 1:12) {
  
  #create FDC data table
  FDC_data_mth <- sim_data %>%
    select(Month, CRO_cumecs) %>%
    filter(Month == a) %>%
    arrange(desc(CRO_cumecs)) %>%
    mutate(rank = dense_rank(desc(CRO_cumecs)))
  
  #number of records
  prob_nrow_mth <- nrow(FDC_data_mth)
  
  #calculate prob dist
  FDC_data_mth <- FDC_data_mth %>%
    mutate(prob = rank/(prob_nrow_mth + 1)*100) %>%
    mutate(id = row_number())
  
  #create FDC table
  FDC_table_mth <- data.frame("Percentage" = c(1:100))
  FDC_table_mth <- FDC_table_mth %>%
    mutate(Tens = as.integer(Percentage/10)*10,
           Ones = Percentage %% 10)
  
  #find matching values
  k <- 1
  FDC_table_mth$index <- 0
  for (k in 1:nrow(FDC_table_mth)) {
    FDC_table_mth$index[k] <- which.min(abs(FDC_data_mth$prob - FDC_table_mth$Percentage[k]))
  }
  
  #match index of matching values
  FDC_table_mth <- FDC_table_mth %>%
    left_join(FDC_data_mth, by = c("index" = "id"))
  
  #pivot
  FDC_table_pv_mth <- FDC_table_mth %>%
    select(Tens, Ones, CRO_cumecs) %>%
    spread(Ones, CRO_cumecs)
  
  #in MLD
  FDC_table_MLD_mth <- FDC_table_mth %>%
    select(Tens, Ones, CRO_cumecs) %>%
    mutate(CRO_MLD = CRO_cumecs*1000/10^6*24*60*60) %>%
    select(Tens, Ones, CRO_MLD) %>%
    spread(Ones, CRO_MLD)
  
  #extract Qn_ex
  #Qn_ex_mth <- FDC_table_mth[Qn, "CRO_cumecs"] #column CRO_cumecs
  Qn_ex_mth <- FDC_table_mth #column CRO_cumecs
  Qn_ex_mth_df <- rbind(Qn_ex_mth_df, data.frame(a, Qn_ex_mth))
  
  
  #name data frames
  assign(paste0("FDC_data_mth", a), FDC_data_mth)
  assign(paste0("FDC_table_mth", a), FDC_table_mth)
  assign(paste0("FDC_table_pv_mth", a), FDC_table_pv_mth)
  assign(paste0("FDC_table_MLD_mth", a), FDC_table_MLD_mth)
  
  list_FDC_mth_cumecs[[a]] <- FDC_table_pv_mth
  list_FDC_mth_MLD[[a]] <- FDC_table_MLD_mth
  

  
}

#combine lists of monthly FDC
#FDC_mth_cumecs = do.call(rbind, list_FDC_mth_cumecs)
#FDC_mth_MLD = do.call(rbind, list_FDC_mth_MLD)

FDC_data_mth_all <- Qn_ex_mth_df %>%
  select(Percentage, Month, CRO_cumecs) %>%
  mutate(CRO_MLD = CRO_cumecs*1000/10^6*24*60*60)

#replace column names

#round values
FDC_data_mth_all$CRO_cumecs <- lapply(FDC_data_mth_all$CRO_cumecs, 
                                round, digits = 3)
FDC_data_mth_all$CRO_MLD <- lapply(FDC_data_mth_all$CRO_MLD, 
                                round, digits = 1)





###############################################################################
#CHARTS
###############################################################################

#FDC
FDC_data2 <- FDC_data %>%
  select(CRO_cumecs, prob)
FDC_data2 %>%
  ggplot(aes(x = prob, y = CRO_cumecs)) +
  geom_line(color = "steelblue", size = 1) +
  geom_segment(aes(x = Qn, y = 0, xend = Qn, yend = Qn_ex),
               linetype = "dashed",
               color = "grey45") +
  geom_segment(aes(x = 0, y = Qn_ex, xend = Qn, yend = Qn_ex),
               linetype = "dashed",
               color = "grey45") +
  annotate("text", x = Qn, y = Qn_ex, 
           label = paste("Q[", Qn,"]==", sprintf("%0.3f", Qn_ex), "~m^3/s==", 
                         sprintf("%0.1f", Qn_ex*1000/10^6*24*60*60), "~MLD"), 
           parse = TRUE,
           vjust = -0.5,
           hjust = 1.0) +
  theme_bw(base_size = 10) +
  scale_x_continuous(name = "Percentage Equaled or Exceeded", 
                     breaks = seq(0, 100, by = 10), 
                     limits = c(0, 100),
                     minor_breaks = NULL,
                     expand = c(0,0)) + #x axis format
  #scale_y_continuous(name = expression(paste("Flow (", m^3, " )")), 
  #                   #breaks = seq(0, 100, by = 10), 
  #                   minor_breaks = NULL,
  #                   expand = c(0,0)) + #y axis format
  scale_y_log10(name = expression(paste("Flow (", m^3, "/s)")), 
                #breaks = seq(from=0, to=10, by = 1), 
                minor_breaks = NULL,
                sec.axis = sec_axis(~.*1000/10^6*24*60*60, 
                                    name = "Flow (MLD)")) +
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12))
#print last plot to file
ggsave(paste0(filename2, "_FDC.jpg"), dpi = 300)


###########################




###########################

#bar chart, annual rainfall
#Total Annual Precipitation
precip_sum_yr %>%
  ggplot(aes(x = Year, y = sum_precip)) +
  geom_bar(stat = "identity", fill = "cornflowerblue") +
  #geom_text(aes(label=sum_precip), position=position_dodge(width=0.9), vjust=-0.3, size=3.5) + #label data
  theme_bw(base_size = 10) +
  scale_x_continuous(name= "Year", 
                     breaks = seq(min_year, max_year, by = 1), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name= "Annual precipitation (mm)",
                     #breaks = seq(0, 3500, by = 500), 
                     minor_breaks = NULL) + #y axis format
  geom_hline(aes(yintercept = mean(sum_precip)), 
             color="black", 
             alpha=0.3, 
             size=1) + #avg line
  geom_text(aes(min_year,
                mean(sum_precip),
                label = paste("Average = ", sprintf("%0.0f", mean(sum_precip)), "mm"), 
                vjust = -0.5, hjust = 0), 
            size=3.5, 
            family="Roboto Light", 
            fontface = 1, 
            color="grey20") + #avg line label
  theme(text=element_text(family="Roboto", color="grey20"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = angle_precip, hjust = 0.5))
#print last plot to file
ggsave(paste0(filename2, "_precip_annual.jpg"), dpi = 300)


#bar chart, monthly rainfall
#Long-term Average Monthly Precipitation
precip_sum_mth %>%
  ggplot(aes(x = Month, y = mth_precip)) +
  geom_bar(stat = "identity", fill = "cornflowerblue") +
  theme_bw(base_size = 10) +
  scale_x_continuous(name = "Month",
                     breaks = seq(1, 12, by = 1), 
                     minor_breaks = NULL, 
                     labels=month.abb) + #x axis format
  scale_y_continuous(name = "Monthly precipitation (mm)",
                     #breaks = seq(0, 350, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20"),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_precip_mth.jpg"), dpi = 300)

###########################

#bar chart, annual runoff (mm)
#Total Annual Runoff (mm)
CRO_mm_sum_yr %>%
  ggplot(aes(x = Year, y = sum_CRO_mm)) +
  geom_bar(stat = "identity", fill = "turquoise3") +
  #geom_text(aes(label=sum_precip), position=position_dodge(width=0.9), vjust=-0.3, size=3.5) + #label data
  theme_bw(base_size = 10) +
  scale_x_continuous(name= "Year", 
                     breaks = seq(min_year, max_year, by = 1), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name= "Annual Runoff (mm)",
                     #breaks = seq(0, 3500, by = 500), 
                     minor_breaks = NULL) + #y axis format
  geom_hline(aes(yintercept = mean(sum_CRO_mm)), 
             color="black", 
             alpha=0.3, 
             size=1) + #avg line
  geom_text(aes(min_year,
                mean(sum_CRO_mm),
                label = paste("Average = ", sprintf("%0.0f", mean(sum_CRO_mm)), "mm"), 
                vjust = -0.5, hjust = 0), 
            size=3.5, 
            family="Roboto Light", 
            fontface = 1, 
            color="grey20") + #avg line label
  theme(text=element_text(family="Roboto", color="grey20"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = angle_precip, hjust = 0.5))
#print last plot to file
ggsave(paste0(filename2, "_CRO_mm_annual.jpg"), dpi = 300)


#bar chart, monthly runoff (mm)
#Long-term Average Monthly runoff (mm)
CRO_mm_sum_mth %>%
  ggplot(aes(x = Month, y = mth_CRO_mm)) +
  geom_bar(stat = "identity", fill = "turquoise3") +
  theme_bw(base_size = 10) +
  scale_x_continuous(name = "Month",
                     breaks = seq(1, 12, by = 1), 
                     minor_breaks = NULL, 
                     labels=month.abb) + #x axis format
  scale_y_continuous(name = "Monthly Runoff (mm)",
                     #breaks = seq(0, 350, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20"),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_CRO_mm_mth.jpg"), dpi = 300)

###########################

#bar chart, annual runoff (m3/s)
#Total Annual Runoff (m3/s)
avg_CRO_cumecs_yr <- mean(CRO_cumecs_sum_yr$sum_CRO_cumecs)
CRO_cumecs_sum_yr %>%
  ggplot(aes(x = Year, y = sum_CRO_cumecs)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  theme_bw(base_size = 10) +
  scale_x_continuous(name= "Year", 
                     breaks = seq(min_year, max_year, by = 1), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name= expression(paste("Annual Runoff (", m^3, "/s )")),
                     #breaks = seq(0, 3500, by = 500), 
                     minor_breaks = NULL) + #y axis format
  geom_hline(aes(yintercept = avg_CRO_cumecs_yr), 
             color="black", 
             alpha=0.3, 
             size=1) + #avg line
  annotate("text", x = min_year, y = avg_CRO_cumecs_yr, 
           label = paste("Average==", sprintf("%0.0f", avg_CRO_cumecs_yr), "~m^3/s"), 
           parse = TRUE,
           vjust = -0.2,
           hjust = 0) +
  theme(text=element_text(family="Roboto", color="grey20"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = angle_precip, hjust = 0.5))
#print last plot to file
ggsave(paste0(filename2, "_CRO_cumecs_annual.jpg"), dpi = 300)


#bar chart, monthly runoff (m3/s)
#Long-term Average Monthly runoff (m3/s)
CRO_cumecs_sum_mth %>%
  ggplot(aes(x = Month, y = mth_CRO_cumecs)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  theme_bw(base_size = 10) +
  scale_x_continuous(name = "Month",
                     breaks = seq(1, 12, by = 1), 
                     minor_breaks = NULL, 
                     labels=month.abb) + #x axis format
  scale_y_continuous(name = expression(paste("Monthly Runoff (", m^3, "/s )")),
                     #breaks = seq(0, 350, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20"),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_CRO_cumecs_mth.jpg"), dpi = 300)

###############################################################################
#OUTPUT
###############################################################################

#rounding
ev_data$PE <- lapply(ev_data$PE, 
                     round, digits = 2)
CRO_mm_sum_yr$sum_CRO_mm <- lapply(CRO_mm_sum_yr$sum_CRO_mm, 
                                   round, digits = 1)
CRO_mm_sum_mth$mth_CRO_mm <- lapply(CRO_mm_sum_mth$mth_CRO_mm, 
                                   round, digits = 1)
CRO_cumecs_sum_yr$sum_CRO_cumecs <- lapply(CRO_cumecs_sum_yr$sum_CRO_cumecs, 
                                   round, digits = 1)
CRO_cumecs_sum_mth$mth_CRO_cumecs <- lapply(CRO_cumecs_sum_mth$mth_CRO_cumecs, 
                                    round, digits = 1)
FDC_table_pv %>% mutate_at(2:11, funs(round(., 2)))
FDC_table_MLD %>% mutate_at(2:11, funs(round(., 1)))



#replace column names in dataframes
ev_data <- ev_data %>% 
  rename(EV = Eto,
         Days = days_mth)

###########################
#summary page
input_param <- c('Basin',
                 'Precipitation Station', 
                 'Length of record (year)', 
                 'Evaporation Station',
                 'Catchment Area (sqkm)',
                 'Water Holding Capacity (mm)',
                 'Initial Soil Moisture (mm)',
                 'Initial available water for runoff (mm)',
                 'Soil moisture retention (a)',
                 'Soil moisture retention (b)',
                 'Recession Constant (K)',
                 'Surface runoff rate (fs)',
                 'Baseflow (mm/day)',
                 '',
                 'BASEFLOW',
                 'Streamflow Station',
                 'Minimum Discharge (m3/s)',
                 'Catchment Area (sqkm)')
                 
input_value <- c(catch_name,
                 station_no,
                 year_w_data,
                 ev_stn,
                 catch_A,
                 water_holding,
                 initial_soil_moisture,
                 initial_avail_water_runoff,
                 soil_mois_retention_a,
                 soil_mois_retention_b,
                 recession_K,
                 surface_runoff_rate,
                 lapply(Qb, round, digits = 3),
                 '',
                 '',
                 bf_stn,
                 bf_Q_min,
                 bf_catch_A)
input_page <- data.frame(cbind(input_param, input_value))


#list all dataframe
list_worksheet <- list("Input" = input_page,
                       "EV" = ev_data,
                       "Calc" = sim_data[c(1:2, 7:26)],
                       "Prob" = FDC_data2,
                       "FDC" = FDC_table[c(1, 7)],
                       "Precip" = precip_sum_yr, 
                       "CRO_mm" = CRO_mm_sum_yr, 
                       "CRO_cumecs" = CRO_cumecs_sum_yr)


# Create a blank workbook
wb <- createWorkbook()

# Loop through the list of split tables as well as their names
#   and add each one as a sheet to the workbook
Map(function(data, name){
  addWorksheet(wb, name)
  writeData(wb, name, data)
}, list_worksheet, names(list_worksheet))

## set col widths
setColWidths(wb, 1, cols = 1:2, widths = "auto")


###########################
#insert data frames
writeData(wb, "Input", station_name, 
          startRow = 3, startCol = 3)
precip_sum_yr_row <- nrow(precip_sum_yr)
writeData(wb, "Precip", precip_sum_mth, 
          startRow = precip_sum_yr_row + 3, startCol = 1)
CRO_mm_sum_yr_row <- nrow(CRO_mm_sum_yr)
writeData(wb, "CRO_mm", CRO_mm_sum_mth, 
          startRow = CRO_mm_sum_yr_row + 3, startCol = 1)
CRO_cumecs_sum_yr_row <- nrow(CRO_cumecs_sum_yr)
writeData(wb, "CRO_cumecs", CRO_cumecs_sum_mth, 
          startRow = CRO_cumecs_sum_yr_row + 3, startCol = 1)

#add new worksheet
addWorksheet(wb, "Table")
writeData(wb, "Table", "Exceedance Frequency (%) of Flow (m3/s)", 
          startRow = 1, startCol = 1)
writeData(wb, "Table", FDC_table_pv, 
          startRow = 2, startCol = 1)
#number and table border formatting
s_cumecs <- createStyle(numFmt = "0.000", border= "TopBottomLeftRight", borderColour = "gray48")
addStyle(wb, "Table", style = s_cumecs, rows = 3:13, cols = 2:11, gridExpand = TRUE)
s_cumecs2 <- createStyle(border= "TopBottomLeftRight", borderColour = "gray48")
addStyle(wb, "Table", style = s_cumecs2, rows = 2:13, cols = 1, gridExpand = TRUE)
addStyle(wb, "Table", style = s_cumecs2, rows = 2, cols = 1:11, gridExpand = TRUE)


writeData(wb, "Table", "Exceedance Frequency (%) of Flow (MLD)", 
          startRow = 15, startCol = 1)
writeData(wb, "Table", FDC_table_MLD, 
          startRow = 16, startCol = 1)
#number and table border formatting
s_mld <- createStyle(numFmt = "0.0", border= "TopBottomLeftRight", borderColour = "gray48")
addStyle(wb, "Table", style = s_mld, rows = 17:27, cols = 2:11, gridExpand = TRUE)
s_mld2 <- createStyle(border= "TopBottomLeftRight", borderColour = "gray48")
addStyle(wb, "Table", style = s_mld2, rows = 16:27, cols = 1, gridExpand = TRUE)
addStyle(wb, "Table", style = s_mld2, rows = 16, cols = 1:11, gridExpand = TRUE)



#add new worksheet
addWorksheet(wb, "Monthly")
writeData(wb, "Monthly", "FDC for each month", 
          startRow = 1, startCol = 1)
writeData(wb, "Monthly", FDC_data_mth_all, 
          startRow = 2, startCol = 1)




#insert image
insertImage(wb, "FDC", paste0(filename2, "_FDC.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Precip", paste0(filename2, "_precip_annual.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Precip", paste0(filename2, "_precip_mth.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 27, startCol = 5, units = "cm")
insertImage(wb, "CRO_mm", paste0(filename2, "_CRO_mm_annual.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "CRO_mm", paste0(filename2, "_CRO_mm_mth.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 27, startCol = 5, units = "cm")
insertImage(wb, "CRO_cumecs", paste0(filename2, "_CRO_cumecs_annual.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "CRO_cumecs", paste0(filename2, "_CRO_cumecs_mth.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 27, startCol = 5, units = "cm")
insertImage(wb, "Table", paste0(filename2, "_FDC.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 29, startCol = 1, units = "cm")


###########################

# Save workbook to working directory
saveWorkbook(wb, file = paste0(filename2, "_output.xlsx"), 
             overwrite = TRUE)
