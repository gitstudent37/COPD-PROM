# Development, Validation, and Clinical Interpretation of a Patient Reported Outcome Measure (PROM) for chronic obstructive pulmonary disease（COPD）

*The related report:*

## *Multicentre development and two-stage validation of a patient-reported outcome measure for chronic obstructive pulmonary disease*

Hangzhi He, Jie Jin, Xiaojuan Hu, Lifang Li, Li Li, Hui Zhao,<sup>#</sup> and Yanbo Zhang<sup>#</sup> 

<sup>#</sup>Co-corresponding authors

<p align="center"><img src="/GraphicAbstract.jpg"></p>

## *Content*

* **development** — analysis scripts and demo datasets used in the development stage
  * **CTT.Rmd** — script for classical test theory (CTT) analysis
  * **IRT.Rmd** — script for item response theory (IRT) analysis
  * **ItemResponseCurve.py** — script for generating item response curves
  * **data_demo.csv** — demo dataset
* **interpretaion** — analysis scripts and demo datasets used in the interpretation stage
  * **PrognosticModel.R** — script for constructing prognostic models and validating MCIDs
  * **data_prognosis_demo.xlsx** — demo dataset for prognostic analyses
  * **data_scores_demo.xlsx** — demo dataset for MCID calculation
  * **mcidCalculation.R** — script for calculating MCIDs
* **validation** — analysis scripts and demo datasets used in the validation stage
  * **data_Corr_demo** — demo datasets for concurrent validity analyses
    * **PRO_demo.xlsx** — demo dataset of PROM scores
    * **SGRQ_demo.xlsx** — demo dataset of SGRQ scores
  * **data_heatmap_demo** — demo datasets for heatmap creation
    * **cor_matrix_demo.csv** — demo correlation matrix
    * **p_matrix_demo.csv** — demo p-value matrix
  * **CFA.R** — script for confirmatory factor analysis
  * **CorrelationAnalysis.R** — script for correlation analysis
  * **Cronbach.R** — script for calculating Cronbach’s α
  * **Heatmap.py** — script for generating heatmaps
  * **SplitHalf.R** — script for calculating split-half reliability
  * **data_demo.xlsx** — demo dataset for validation analyses



<u>**ATTENTION:**</u> The sample data were generated using population-level statistical characteristics and are provided solely for educational and illustrative purposes. They must not be used as the basis for research or reporting.



## *Get started*

To reproduce the analyses presented in the report,  follow these steps:

* You can download the ZIP file via the code repository, clicking the **Code** button, and selecting **Download ZIP**.

* Alternatively, clone the repository by running the following command in Git:

   ``````bash
   git clone https://github.com/gitstudent37/COPD-PROM.git
   ``````

* It is recommended to run the scripts using IDEs of R (e.g., [RStudio](https://posit.co/download/rstudio-desktop/ "https://posit.co/download/rstudio-desktop/")) or Python (e.g., [PyCharm](https://www.jetbrains.com/pycharm/ "https://www.jetbrains.com/pycharm/"))

