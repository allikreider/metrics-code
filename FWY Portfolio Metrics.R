#install.packages("readxl")
library(readxl)
#install.packages("lubridate")
library(lubridate)
#install.packages("basictabler")
library(basictabler)
#install.packages("dplyr")
library(dplyr)
#install.packages("openxlsx")
library(openxlsx)
#install.packages("writexl")
library(writexl)
#install.packages("scales")
library(scales)
#install.packages("tidyr")
library(tidyr)

portfolio <- read_excel("/Users/2020macbookpro/Downloads/On Course Capital LLC/Portfolio Metrics/FWY PORTFOLIO.xlsx")
#imports updated portfolio into R

portfolio <- portfolio %>%
  mutate(booking_year = year(portfolio$booking_date)) %>%
  mutate(expiration_year = year(portfolio$last_payment_date))
#adds variables for year booked and year expiring

schedule <- portfolio %>%
  filter(percent_of_schedule == 1)
#creates dataset of only schedules within the portfolio

asset <- portfolio %>%
  filter(!is.na(asset_description))
#creates dataset of all assets within portfolio

schedules_per_customer_fwy <- schedule %>%
  filter(original_lease_holder == "FWY") %>%
  group_by(name) %>% 
  summarise(schedules = n()) 
#find schedules per fwy customer
volume_per_customer_fwy <- schedule %>%
  filter(original_lease_holder == "FWY") %>%
  group_by(name) %>%
  summarise(volume = sum(full_schedule_funding))
#find volume per fwy customer
net_invest_25_per_customer_fwy <- schedule %>%
  filter(original_lease_holder == "FWY") %>%
  group_by(name) %>%
  drop_na(march_2025_buyout) %>%
  summarise(net_investment_2025 = sum(march_2025_buyout))
#find net investment 2025 per fwy customer
net_invest_26_per_customer_fwy <- schedule %>%
  filter(original_lease_holder == "FWY") %>%
  group_by(name) %>%
  drop_na(march_2026_buyout) %>%
  summarise(net_investment_2026 = sum(march_2026_buyout))
#find net investment 2026 per fwy customer
assets_per_customer_fwy <- asset %>%
  filter(original_lease_holder == "FWY") %>%
  group_by(name) %>%
  summarise(assets = n())
#find assets per fwy customer

table1_fwy <- merge(schedules_per_customer_fwy,volume_per_customer_fwy)
table1_fwy <- merge(table1_fwy, net_invest_25_per_customer_fwy)
table1_fwy <- merge(table1_fwy, net_invest_26_per_customer_fwy)
table1_fwy <- merge(table1_fwy, assets_per_customer_fwy)
#merge all fwy datasets

total_fwy <- table1_fwy %>%
  summarise(name = "FWY", 
            schedules = sum(schedules), 
            volume = sum(volume), 
            net_investment_2025 = sum(net_investment_2025), 
            net_investment_2026 = sum(net_investment_2026), 
            assets = sum(assets))
#creates top line total for fwy

table1_fwy <- bind_rows(total_fwy, table1_fwy)
#adds total row to top of fwy data

schedules_per_customer_nuco <- schedule %>%
  filter(original_lease_holder == "NUCOPSS") %>%
  group_by(name) %>% 
  summarise(schedules = n()) 
#find schedules per nuco customer
volume_per_customer_nuco <- schedule %>%
  filter(original_lease_holder == "NUCOPSS") %>%
  group_by(name) %>%
  summarise(volume = sum(full_schedule_funding))
#find volume per nuco customer
net_invest_25_per_customer_nuco <- schedule %>%
  filter(original_lease_holder == "NUCOPSS") %>%
  group_by(name) %>%
  drop_na(march_2025_buyout) %>%
  summarise(net_investment_2025 = sum(march_2025_buyout))
#find net investment 2025 per nuco customer
net_invest_26_per_customer_nuco <- schedule %>%
  filter(original_lease_holder == "NUCOPSS") %>%
  group_by(name) %>%
  drop_na(march_2026_buyout) %>%
  summarise(net_investment_2026 = sum(march_2026_buyout))
#find net investment 2026 per nuco customer
assets_per_customer_nuco <- asset %>%
  filter(original_lease_holder == "NUCOPSS") %>%
  group_by(name) %>%
  summarise(assets = n())
#find assets per nuco customer

table1_nuco <- merge(schedules_per_customer_nuco,volume_per_customer_nuco)
table1_nuco <- merge(table1_nuco, net_invest_25_per_customer_nuco)
table1_nuco <- merge(table1_nuco, net_invest_26_per_customer_nuco)
table1_nuco <- merge(table1_nuco, assets_per_customer_nuco)
#merge all nuco datasets

total_nuco <- table1_nuco %>%
  summarise(name = "NUCOPSS", 
            schedules = sum(schedules), 
            volume = sum(volume), 
            net_investment_2025 = sum(net_investment_2025), 
            net_investment_2026 = sum(net_investment_2026), 
            assets = sum(assets))
#creates top line total for nuco

table1_nuco <- bind_rows(total_nuco, table1_nuco)
#adds total row to top of nuco data

schedules_per_customer_tsp <- schedule %>%
  filter(original_lease_holder == "TSP") %>%
  group_by(name) %>% 
  summarise(schedules = n()) 
#find schedules per tsp customer
volume_per_customer_tsp <- schedule %>%
  filter(original_lease_holder == "TSP") %>%
  group_by(name) %>%
  summarise(volume = sum(full_schedule_funding))
#find volume per tsp customer
net_invest_25_per_customer_tsp <- schedule %>%
  filter(original_lease_holder == "TSP") %>%
  group_by(name) %>%
  drop_na(march_2025_buyout) %>%
  summarise(net_investment_2025 = sum(march_2025_buyout))
#find net investment 2025 per tsp customer
net_invest_26_per_customer_tsp <- schedule %>%
  filter(original_lease_holder == "TSP") %>%
  group_by(name) %>%
  drop_na(march_2026_buyout) %>%
  summarise(net_investment_2026 = sum(march_2026_buyout))
#find net investment 2026 per tsp customer
assets_per_customer_tsp <- asset %>%
  filter(original_lease_holder == "TSP") %>%
  group_by(name) %>%
  summarise(assets = n())
#find assets per tsp customer

table1_tsp <- merge(schedules_per_customer_tsp,volume_per_customer_tsp)
table1_tsp <- merge(table1_tsp, net_invest_25_per_customer_tsp)
table1_tsp <- merge(table1_tsp, net_invest_26_per_customer_tsp)
table1_tsp <- merge(table1_tsp, assets_per_customer_tsp)
#merge all tsp datasets

total_tsp <- table1_tsp %>%
  summarise(name = "TSP", 
            schedules = sum(schedules), 
            volume = sum(volume), 
            net_investment_2025 = sum(net_investment_2025), 
            net_investment_2026 = sum(net_investment_2026), 
            assets = sum(assets))
#creates top line total for tsp

table1_tsp <- bind_rows(total_tsp, table1_tsp)
#adds total row to top of tsp data

table1 <- bind_rows(table1_fwy, table1_nuco, table1_tsp)

total_customers <- paste("Total Customers:", (length(unique(schedule$name))), "\n")
#counts customers and prints in correct format
table1 <- table1 %>%
      bind_rows(summarise(., across(where(is.numeric), sum),
                          across(where(is.character), ~total_customers)))
#adds total row to table1

table1 <- table1 %>%
  rename("Customer Name" = name, 
         "Schedules per Customer" = schedules, 
         "Volume per Customer" = volume, 
         "Net Investment (March 2025)" = net_investment_2025, 
         "Net Investment (March 2026)" = net_investment_2026, 
         "Assets per Customer" = assets)
#rename columns

schedules_booked <- schedule %>%
  group_by(booking_year) %>%
  summarise(schedules = n())
#finds schedules booked per year
assets_booked <- asset %>%
  group_by(booking_year) %>%
  summarise(assets = n())
#finds assets booked per year
volume_booked <- schedule %>%
  group_by(booking_year) %>%
  summarise(volume_per_year = sum(full_schedule_funding))
#finds volume booked per year
residuals_booked <- schedule %>%
  group_by(booking_year) %>%
  summarise(residuals_per_year = sum(full_residual))
#finds residuals booked per year

table2 <- merge(schedules_booked, assets_booked)
table2 <- merge(table2, volume_booked)
table2 <- merge(table2, residuals_booked)
#merge all datasets
total_booked <- data.frame(booking_year = c("Totals:"),
                           schedules = sum(table2$schedules),
                           assets = sum(table2$assets),
                           volume_per_year = sum(table2$volume_per_year), 
                           residuals_per_year = sum(table2$residuals_per_year))
#make total row for table2

table2 <- rbind(table2, total_booked)
#add total row to table2

table2 <- table2 %>%
  rename("Booking Year" = booking_year, 
         "Schedules Booked" = schedules, 
         "Assets Booked" = assets, 
         "Volume Booked" = volume_per_year, 
         "Residuals Booked" = residuals_per_year)
#rename columns

schedules_expiring <- schedule %>%
  group_by(expiration_year) %>%
  summarise(schedules = n())
#finds schedules expiring per year
assets_expiring <- asset %>%
  group_by(expiration_year) %>%
  summarise(assets = n())
#finds assets expiring per year
volume_expiring <- schedule %>%
  group_by(expiration_year) %>%
  summarise(volume_per_year = sum(full_schedule_funding))
#finds volume expiring per year
residuals_expiring <- schedule %>%
  group_by(expiration_year) %>%
  summarise(residuals_per_year = sum(full_residual))
#finds residuals expiring per year

table3 <- merge(schedules_expiring, assets_expiring)
table3 <- merge(table3, volume_expiring)
table3 <- merge(table3, residuals_expiring)
#merge all datasets
total_expiring <- data.frame(expiration_year = c("Totals:"),
                           schedules = sum(table3$schedules),
                           assets = sum(table3$assets),
                           volume_per_year = sum(table3$volume_per_year), 
                           residuals_per_year = sum(table3$residuals_per_year))
#make total row for table3

table3 <- rbind(table3, total_expiring)
#add total row to table3

table3 <- table3 %>%
  rename("Expiration Year" = expiration_year, 
         "Schedules Expiring" = schedules, 
         "Assets Expiring" = assets, 
         "Volume Expiring" = volume_per_year, 
         "Residuals Expiring" = residuals_per_year)
#rename columns

schedule_lease <- schedule %>%
  group_by(lease_type) %>%
  summarise(schedules = n())
#finds number of schedules by lease type
schedule_lease <- schedule_lease %>%
  mutate(schedule_percentage = schedules/sum(schedules))
#finds percentage of schedules by lease type
asset_lease <- asset %>%
  group_by(lease_type) %>%
  summarise(assets = n())
#finds number of assets by lease type
asset_lease <- asset_lease %>%
  mutate(asset_percentage = assets/sum(assets))
#finds percentage of volume by lease type
volume_lease <- schedule %>%
  group_by(lease_type) %>%
  summarise(volume = sum(full_schedule_funding))
#finds number of volume by lease type
volume_lease <- volume_lease %>%
  mutate(volume_percentage = volume/sum(volume))
#finds percentage of volume by lease type

table4 <- merge(schedule_lease, asset_lease)
table4 <- merge(table4, volume_lease)
#create table4

table4 <- table4 %>%
  rename("Lease Type" = lease_type, 
         "Schedules" = schedules, 
         "Percentage of Schedules" = schedule_percentage, 
         "Assets" = assets, 
         "Percentage of Assets" = asset_percentage,
         "Volume" = volume, 
         "Percentage of Volume" = volume_percentage)
#rename columns

table5 <- asset %>%
  group_by(asset_type) %>%
  summarise(assets_by_type = n())
#finds number of assets by asset type
table5 <- table5 %>%
  mutate(percentage = assets_by_type/sum(assets_by_type))
#finds percentage of assets by asset type

table5 <- table5 %>%
  rename("Asset Type" = asset_type, 
         "Number of Assets" = assets_by_type, 
         "Percentage of Assets" = percentage)
#rename columns

table6 <- schedule %>%
  group_by(asset_type) %>%
  summarise(schedules_by_type = n())
#finds number of schedules by schedule type
table6 <- table6 %>%
  mutate(percentage = schedules_by_type/sum(schedules_by_type))
#finds percentage of schedules by schedule type
table6 <- table6 %>%
  rename("Schedule Type" = asset_type, 
         "Number of Schedules" = schedules_by_type, 
         "Percentage of Schedules" = percentage)
#rename columns

schedule_segment <- schedule %>%
  group_by(customer_segment) %>%
  summarise(schedules_by_segment = n())
#finds number of schedules by customer segment
schedule_segment <- schedule_segment %>%
  mutate(percentage = schedules_by_segment/sum(schedules_by_segment))
#finds percentage of schedules by customer segment
volume_segment <- schedule %>%
  group_by(customer_segment) %>%
  summarise(volume = sum(full_schedule_funding))
#finds volume by customer segment

table7 <- merge(schedule_segment, volume_segment)
#creates table7

table7 <- table7 %>%
  rename("Customer Segment" = customer_segment, 
         "Number of Schedules" = schedules_by_segment, 
         "Percentage of Schedules" = percentage, 
         "Volume" = volume)
#rename columns

average_spo_residual <- schedule %>%
  group_by(account_number) %>%
  filter(lease_type == "SPO") %>%
  summarise(average_spo_residual_percentage = sum(full_residual)/sum(full_schedule_funding))
#calculates average spo residual as a percentage of full cost

average_fmv_residual <- schedule %>%
  group_by(account_number) %>%
  filter(lease_type == "FMV") %>%
  summarise(average_fmv_residual_percentage = sum(full_residual)/sum(full_schedule_funding))
#calculates average fmv residual as a percentage of full cost

average_assets_customer <- asset %>%
  group_by(account_number) %>%
  summarise(average_assets_per_customer = n()/length(unique(portfolio$name)))
#calculates average number of assets per customer

average_assets_schedule <- asset %>%
  group_by(account_number) %>%
  summarise(average_assets_per_schedule = n()/length(unique(portfolio$contract_number)))
#calculates average number of assets per schedule

average_asset <- asset %>%
  group_by(account_number) %>%
  filter(!is.na(equipment_cost)) %>%
  summarise(average_asset_cost = mean(equipment_cost))
#calculates average asset cost

average_schedules_customer <- schedule %>%
  group_by(account_number) %>%
  summarise(average_schedules_per_customer = n()/length(unique(portfolio$name)))
#calculates the average number of schedules per customer

average_customer_volume <- schedule %>%
  group_by(account_number) %>%
  summarise(average_volume_per_customer = sum(full_schedule_funding)/length(unique(portfolio$name)))
#calculates average volume per customer

average_schedule_volume <- schedule %>%
  group_by(account_number) %>%
  summarise(average_volume_per_schedule = mean(full_schedule_funding))
#calculates the average volume per schedule

customers_20 <- schedule %>%
  group_by(name) %>%
  filter(booking_year == 2020) %>%
  summarise(schedules = n())
#finds customers with schedules in 2020
customers_21 <- schedule %>%
  group_by(name) %>%
  filter(booking_year == 2021) %>%
  summarise(schedules = n())
#finds customers with schedules in 2021
customers_20_21 <- customers_20 %>%
  filter(name %in% customers_21$name) %>%
  summarise(repeat_rate_20_21 = n()/length(customers_21$name))
#finds repeat rate for 2020-2021

customers_22 <- schedule %>%
  group_by(name) %>%
  filter(booking_year == 2022) %>%
  summarise(schedules = n())
#finds customers with schedules in 2022
customers_21_22 <- customers_21 %>%
  filter(name %in% customers_22$name) %>%
  summarise(repeat_rate_21_22 = n()/length(customers_22$name))
#finds repeat rate for 2021-2022

customers_23 <- schedule %>%
  group_by(name) %>%
  filter(booking_year == 2023) %>%
  summarise(schedules = n())
#finds customers with schedules in 2023
customers_22_23 <- customers_22 %>%
  filter(name %in% customers_23$name) %>%
  summarise(repeat_rate_22_23 = n()/length(customers_23$name))
#finds repeat rate for 2023-2024

customers_24 <- schedule %>%
  group_by(name) %>%
  filter(booking_year == 2024) %>%
  summarise(schedules = n())
#finds customers with schedules in 2024
customers_23_24 <- customers_23 %>%
  filter(name %in% customers_24$name) %>%
  summarise(repeat_rate_23_24 = n()/length(customers_23$name))
#finds repeat rate for 2023-2024

#for new repeat rates, copy previous lines and change dataset names and filters to accomodate correct years

table8 <- merge(average_spo_residual, average_fmv_residual)
table8 <- merge(table8, average_assets_customer)
table8 <- merge(table8, average_assets_schedule)
table8 <- merge(table8, average_asset)
table8 <- merge(table8, average_schedules_customer)
table8 <- merge(table8, average_customer_volume)
table8 <- merge(table8, average_schedule_volume)
table8 <- merge(table8, customers_20_21)
table8 <- merge(table8, customers_21_22)
table8 <- merge(table8, customers_22_23)
table8 <- merge(table8, customers_23_24)
#create table8

table8 <- table8 %>%
  select(-account_number)
#remove accountnumber column

table8 <- table8 %>%
  rename("Average SPO Residual Percentage" = average_spo_residual_percentage, 
         "Average FMV Residual Percentage" = average_fmv_residual_percentage, 
         "Average Assets per Customer" = average_assets_per_customer, 
         "Average Assets per Schedule" = average_assets_per_schedule, 
         "Average Asset Cost" = average_asset_cost, 
         "Average Schedules per Customer" = average_schedules_per_customer, 
         "Average Volume per Customer" = average_volume_per_customer, 
         "Average Volume per Schedule" = average_volume_per_schedule, 
         "Repeat Rate 2020-21" = repeat_rate_20_21,
         "Repeat Rate 2021-22" = repeat_rate_21_22,
         "Repeat Rate 2022-23" = repeat_rate_22_23, 
         "Repeat Rate 2023-24" = repeat_rate_23_24)
#rename columns

table9 <- schedule %>%
  filter(expiration_year == 2024) %>%
  select(contract_number, name, dba, last_payment_date) %>%
  arrange(last_payment_date)
#all schedules expiring 2024

table9 <- table9 %>%
  rename("Contract Number" = contract_number, 
         "Customer Name" = name, 
         "DBA" = dba, 
         "Last Payment Date" = last_payment_date)
#rename columns

table10 <- schedule %>%
  filter(expiration_year == 2025) %>%
  select(contract_number, name, dba, last_payment_date) %>%
  arrange(last_payment_date)
#all schedules expiring 2025

table10 <- table10 %>%
  rename("Contract Number" = contract_number, 
         "Customer Name" = name, 
         "DBA" = dba, 
         "Last Payment Date" = last_payment_date)
#rename columns

data_list <- list(Sheet1 = table1, 
                  Sheet2 = table2, 
                  Sheet3 = table3, 
                  Sheet4 = table4, 
                  Sheet5 = table5, 
                  Sheet6 = table6, 
                  Sheet7 = table7, 
                  Sheet8 = table8, 
                  Sheet9 = table9, 
                  Sheet10 = table10)

write_xlsx(data_list, path = "FWY Metrics.xlsx")

wb <- createWorkbook()

addWorksheet(wb, "Sheet1")
writeData(wb, sheet = "Sheet1", table1)
addWorksheet(wb, "Sheet2")
writeData(wb, sheet = "Sheet2", table2)
addWorksheet(wb, "Sheet3")
writeData(wb, sheet = "Sheet3", table3)
addWorksheet(wb, "Sheet4")
writeData(wb, sheet = "Sheet4", table4)
addWorksheet(wb, "Sheet5")
writeData(wb, sheet = "Sheet5", table5)
addWorksheet(wb, "Sheet6")
writeData(wb, sheet = "Sheet6", table6)
addWorksheet(wb, "Sheet7")
writeData(wb, sheet = "Sheet7", table7)
addWorksheet(wb, "Sheet8")
writeData(wb, sheet = "Sheet8", table8)
addWorksheet(wb, "Sheet9")
writeData(wb, sheet = "Sheet9", table9)
addWorksheet(wb, "Sheet10")
writeData(wb, sheet = "Sheet10", table10)

dollar_format <- createStyle(numFmt = "$#,##0.00")
percentage_format <- createStyle(numFmt = "0.00%")
date_format <- createStyle(numFmt = "yyyy-mm-dd")
decimal_format <- createStyle(numFmt = "0.00")
#formats numbers correctly

addStyle(wb, sheet = "Sheet1", style = dollar_format, cols = 3:5, rows = 2:(nrow(table1) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet2", style = dollar_format, cols = 4:5, rows = 2:(nrow(table2) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet3", style = dollar_format, cols = 4:5, rows = 2:(nrow(table3) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet4", style = percentage_format, cols = 3, rows = 2:(nrow(table4) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet4", style = percentage_format, cols = 5, rows = 2:(nrow(table4) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet4", style = percentage_format, cols = 7, rows = 2:(nrow(table4) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet4", style = dollar_format, cols = 6, rows = 2:(nrow(table4) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet5", style = percentage_format, cols = 3, rows = 2:(nrow(table5) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet6", style = percentage_format, cols = 3, rows = 2:(nrow(table6) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet7", style = percentage_format, cols = 3, rows = 2:(nrow(table7) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet7", style = dollar_format, cols = 4, rows = 2:(nrow(table7) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = percentage_format, cols = 1:2, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = decimal_format, cols = 3:4, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = dollar_format, cols = 5, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = decimal_format, cols = 6, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = dollar_format, cols = 7:8, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = percentage_format, cols = 9:12, rows = 2:(nrow(table8) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet9", style = date_format, cols = 4, rows = 2:(nrow(table9) + 1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet10", style = date_format, cols = 4, rows = 2:(nrow(table10) + 1), gridExpand = TRUE)
#formats all numbers correctly

header_style <- createStyle(textDecoration = "Bold", 
                            fgFill = "skyblue1", 
                            border = "TopBottom", 
                            borderColour = "skyblue4")
#top row is highlighted, bold, and has a border in darker blue on top and bottom

addStyle(wb, sheet = "Sheet1", style = header_style, rows = 1, cols = 1:length(table1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet1", style = header_style, rows = nrow(table1) + 1, cols = 1:length(table1), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet2", style = header_style, rows = 1, cols = 1:length(table2), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet2", style = header_style, rows = nrow(table2) + 1, cols = 1:length(table2), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet3", style = header_style, rows = 1, cols = 1:length(table3), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet3", style = header_style, rows = nrow(table3) + 1, cols = 1:length(table2), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet4", style = header_style, rows = 1, cols = 1:length(table4), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet5", style = header_style, rows = 1, cols = 1:length(table5), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet6", style = header_style, rows = 1, cols = 1:length(table6), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet7", style = header_style, rows = 1, cols = 1:length(table7), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet8", style = header_style, rows = 1, cols = 1:length(table8), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet9", style = header_style, rows = 1, cols = 1:length(table9), gridExpand = TRUE)
addStyle(wb, sheet = "Sheet10", style = header_style, rows = 1, cols = 1:length(table10), gridExpand = TRUE)
#makes header rows bold font and highlighted

saveWorkbook(wb, file = "FWY Metrics.xlsx", overwrite = TRUE)
#saves workbook
metrics <- loadWorkbook("FWY Metrics.xlsx")
names <- names(metrics)
for (sheet in names) {
  num_cols <- ncol(read.xlsx(metrics, sheet = sheet))
  setColWidths(metrics, sheet = sheet, cols = 1:num_cols, widths = 31)}
#resizes columns in excel file

saveWorkbook(metrics, file = "FWY Metrics.xlsx", overwrite = TRUE)
#resaves workbook
system("open /Users/2020macbookpro/FWY Metrics.xlsx")
#opens workbook

