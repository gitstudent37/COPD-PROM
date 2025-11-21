# [file name]: PrognosticModel.R
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


library(dplyr)
library(broom)
library(ggplot2)
library(pROC)
library(openxlsx)
library(officer)
library(flextable)
library(stringr)

run_logistic_analysis <- function(df0, predictor, doc = NULL, outcome = "outcome", roc_path_prefix = "ROC_") {
  
  # 1. Determination of covariates
  covariates <- c("Sex", "Exposure.of.inhalation.substances",
                  "Smoking.history", "Drinking.history", "mMRC.grades", 
                  "Family.history.of.COPD", "Exacerbation.history", 
                  "cardiovascular.diseases", "arrhythmia", "hypertension", 
                  "bronchiectasis", "obstructive.sleep.apnea", "diabetes", 
                  "anxiety.and.depression", "Brochodilators", "Glucocorticoids", 
                  "Antibiotics", "COVID.19history", "Age")
  
  # Prognostic model formula
  formula_full <- as.formula(paste0(outcome, " ~ ", predictor, " + ", paste(covariates, collapse = " + ")))
  
  # 2. Model fitting
  fit_full <- glm(formula_full, data = df0, family = binomial(link = "logit"))
  
  # 3. Output of OR, 95% CI, and p-value.
  result_full <- broom::tidy(fit_full, conf.int = TRUE, exponentiate = TRUE) %>%
    dplyr::select(term, estimate, conf.low, conf.high, p.value) %>%
    dplyr::rename(OR = estimate, CI_lower = conf.low, CI_upper = conf.high, P_value = p.value)
  
  # 4. ROC drawing
  plot_roc <- function(fit, df0, title, file_path) {
    df0$predicted_prob <- predict(fit, type = "response")
    roc_obj <- pROC::roc(df0[[outcome]], df0$predicted_prob)
    auc_value <- pROC::auc(roc_obj)
    
    roc_data <- data.frame(
      sensitivity = roc_obj$sensitivities,
      specificity = roc_obj$specificities
    )
    roc_data$one_minus_specificity <- 1 - roc_data$specificity
    
    title_clean <- stringr::str_replace_all(title, "_1", "")
    
    p <- ggplot(roc_data, aes(x = one_minus_specificity, y = sensitivity)) +
      geom_line(color = "#2E86AB", linewidth = 1.2) +
      geom_abline(slope = 1, intercept = 0, color = "grey", linetype = "dashed") +
      labs(title = title_clean, 
           x = "1 - Specificity", 
           y = "Sensitivity") +
      theme_classic() +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold", size = 14)
      ) +
      annotate("text", x = 0.6, y = 0.2,
               label = paste0("AUC = ", round(auc_value, 3)),
               size = 5, color = "#A23B72", fontface = "bold") +
      scale_x_continuous(limits = c(0, 1), breaks = seq(0, 1, 0.2)) +
      scale_y_continuous(limits = c(0, 1), breaks = seq(0, 1, 0.2))
    
    # Ensure the directory exists
    dir.create(dirname(file_path), showWarnings = FALSE, recursive = TRUE)
    
    ggsave(filename = file_path, plot = p, width = 6, height = 5)
    
    return(list(plot = p, auc = auc_value, roc_obj = roc_obj))
  }
  
  # ROC Saving
  roc_full <- plot_roc(fit_full, df0, paste0(predictor, " Full Model ROC"), 
                       paste0(roc_path_prefix, predictor, "_full.png"))
  
  # 5. Output for Word
  if (is.null(doc)) {
    doc <- officer::read_docx()
  }
  
  # Formatting the table
  format_flextable <- function(df) {
    ft <- flextable::flextable(df) %>%
      flextable::theme_box() %>%
      flextable::autofit()
    return(ft)
  }
  
  # Add the table to a Word document
  doc <- doc %>% 
    officer::body_add_par(value = paste0(predictor, "Model regression results"), style = "heading 1") %>%
    { 
      # Adding model table
      flextable::body_add_flextable(., value = format_flextable(result_full)) 
    } %>%
    officer::body_add_par(value = "Model ROC curve", style = "heading 2") %>%
    officer::body_add_img(src = paste0(roc_path_prefix, predictor, "_full.png"), width = 6, height = 5) %>%
    officer::body_add_break()
  
  return(list(doc = doc,
              full_model = result_full,
              roc_full = roc_full))
}


# Input data
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
df <- read.xlsx("data_prognosis_demo.xlsx")
df <- df %>%
  rename(COVID.19history = `COVID-19.history`)

# Categorical variables
factor_vars <- c(
  "Sex", "Exposure.of.inhalation.substances", "Smoking.history",
  "Drinking.history", "Family.history.of.COPD", "Exacerbation.history",
  "cardiovascular.diseases", "arrhythmia", "hypertension", "bronchiectasis",
  "anxiety.and.depression", "Brochodilators","Glucocorticoids", "Antibiotics",
  "COVID.19history", "obstructive.sleep.apnea", "diabetes",
  "SPE", "GEN", "IND", "ANX", "DEP", "COG", "IMP", "SUP", "TAD", "ADR", "SAT",
  "PHD", "PSD", "SOD", "THD", "Total","outcome"
)

# Specifying the variable category
df <- df %>% mutate(
    Age = as.numeric(Age),
    mMRC.grades= as.numeric(mMRC.grades),
    across(all_of(factor_vars), as.factor)
  )

#----------------------------
# Predictors Analysis
#----------------------------
predictors <- c("SPE", "GEN", "IND", "ANX", "DEP", "COG", "IMP",
                "SUP", "TAD", "ADR", "SAT", "PHD", "PSD", "SOD", "THD", "Total")

# Saving paths
output_file <- "./SPE_logistic_results.docx"
roc_prefix <- "./ROC_"

# Initialization
doc_all <- officer::read_docx()

# Iterative analysis of predictors
results_list <- list()
for (pred in predictors) {
  tryCatch({
    cat("Analysing:", pred, "\n")
    res <- run_logistic_analysis(df, pred, doc = doc_all, roc_path_prefix = roc_prefix)
    doc_all <- res$doc
    results_list[[pred]] <- res
    cat("Complete:", pred, "\n\n")
  }, error = function(e) {
    cat("Error in analysis of", pred, ":", e$message, "\n\n")
  })
}

# Output
print(doc_all, target = output_file)
cat("Analysis complete! Results saved to:", output_file, "\n")

# [file content end]