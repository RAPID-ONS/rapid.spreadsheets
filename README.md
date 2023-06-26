# RAPID spreadsheets

This package helps create reproducible spreadsheets that meet [Analysis Function accessibility guidelines](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/). It was designed to help government analysts to produce data tables for their Reproducible Analytical Pipelines. 
rapid.spreadsheets package supports creating reference tables that can include:
* Cover sheet
* Contents table 
* Notes table 
* Data tables

Aside from data tables, all sheet types support use of internal and external hyperlinks. As rapid.spreadsheets package was built using openxlsx workbooks it provides users with flexibility to add and modify formatting styles using openxlsx functions.

## Installation instructions

To install the package you can either:

Download the package repository and run:

```{r}
devtools::install("rapid.spreadsheets", build_vignettes = TRUE)
```

or install directly from GitHub:
```{r}
devtools:: install_github("RAPID-ONS/rapid.spreadsheets", build_vignettes = TRUE)
```
## User manual
To view vignettes (tutorial) on how to use the package, run the following code in the console after installation:
```{r}
browseVignettes("rapid.spreadsheets")
```
