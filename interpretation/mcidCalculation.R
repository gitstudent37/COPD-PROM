# [file name]: mcidCalculation.R
# [file content begin]
# 
#   Copyright (C) 2025 SXMU
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
# http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


library(openxlsx)
library(dplyr)
library(writexl)

# Input data
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
df <- read.xlsx("data_scores_demo.xlsx")


##------ Within-patients difference (WD) --------
target_cols <- 4:19

# Filtering patients with poor prognosis
df1 <- df %>% filter(outcome == 1)

# Baseline（time=0）
baseline <- df1 %>% 
  filter(time == 0) %>% 
  select(number, all_of(target_cols))

# Last follow-up (time != 0)
last_visit <- df1 %>%
  filter(time != 0) %>%
  select(number, all_of(target_cols))

# Merging baseline and the last
merged <- baseline %>%
  inner_join(last_visit, by = "number", suffix = c("_base", "_last"))

# WD = | mean(last - base) |
WD_raw <- sapply(colnames(baseline)[-1], function(var){
  base_col <- merged[[paste0(var, "_base")]]
  last_col <- merged[[paste0(var, "_last")]]
  abs(mean(last_col - base_col, na.rm = TRUE))
})


group0 <- df1 %>% filter(time == 0) %>% select(all_of(target_cols))
group1 <- df1 %>% filter(time != 0) %>% select(all_of(target_cols))

result <- data.frame(
  Mean1  = sapply(group0, mean, na.rm = TRUE),
  SD1    = sapply(group0, sd, na.rm = TRUE),
  Mean2  = sapply(group1, mean, na.rm = TRUE),
  SD2    = sapply(group1, sd, na.rm = TRUE),
  WD_abs = WD_raw
)

rownames(result) <- colnames(group0)

print(result)
##-------- Effect size (ES)/0.5ES--------
#  ES
result$ES <- abs(result$Mean2 - result$Mean1)

#  0.5ES 
result$`0.5ES` <- 0.5 * result$SD1

##-------- Standardised response mean (SRM)--------
diff_list <- lapply(colnames(baseline)[-1], function(var){
  merged[[paste0(var, "_last")]] - merged[[paste0(var, "_base")]]
})

names(diff_list) <- colnames(baseline)[-1]

# Mean difference d and standard deviation d_sd for each metric
d <- sapply(diff_list, function(x) mean(x, na.rm = TRUE))
d_sd <- sapply(diff_list, function(x) sd(x, na.rm = TRUE))

# The calculation results added
result$d <- d
result$d_sd <- d_sd
result$SRM <- abs(result$Mean2 - result$Mean1) / result$d_sd
result$α <- c(
  0.878, 0.724, 0.819, 0.835, 0.826, 0.727, 0.826, 0.831,
  0.938, 0.830, 0.912, 0.903, 0.890, 0.710, 0.775, 0.911
) # Cronbach’s α for each metric (artificial input)

##-------- Standard error of measurement (SEM) --------
result$SEM <- result$SD1 * sqrt(1 - result$α)*1.96
##-------- Reliable change index (RCI) --------
result$RCI <- 1.96 * abs(result$Mean2 - result$Mean1) / sqrt(2 * result$SEM^2)

# Output
result <- result %>% select(-c(Mean1, Mean2, SD1, SD2, d, d_sd, α))
result <- data.frame(Metric = rownames(result), result, row.names = NULL)
write_xlsx(result, "MICD.xlsx")

# [file content end]