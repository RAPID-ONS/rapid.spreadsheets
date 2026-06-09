# RAPID spreadsheets

> [!WARNING]
>
> This repository is now maintained in the ONS Digital GitHub organisation.
>
> https://github.com/ONSdigital/rapid.spreadsheets
>
> This repo is archived / no longer actively maintained.  Please direct all issues and pull requests to the new location.

This package helps create reproducible spreadsheets that meet [Analysis Function accessibility guidelines](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/). It was designed to help government analysts to produce data tables for their Reproducible Analytical Pipelines. 
rapid.spreadsheets package supports creating reference tables that can include:
* Cover sheet
* Contents table 
* Notes table 
* Data tables

Aside from data tables, all sheet types support use of internal and external hyperlinks. As rapid.spreadsheets package was built using `openxlsx` workbooks it provides users with flexibility to add and modify formatting styles using `openxlsx` functions.

## Installation instructions

To install the package you can either:

Download the package repository and run:

```{r}
devtools::install("rapid.spreadsheets", dependencies = TRUE, build_vignettes = TRUE)
```

or install directly from GitHub:
```{r}
devtools::install_github("RAPID-ONS/rapid.spreadsheets", dependencies = TRUE, build_vignettes = TRUE)
```
## User manual
To view vignettes (tutorial) on how to use the package, run the following code in the console after installation:
```{r}
browseVignettes("rapid.spreadsheets")
```

## Maintainers:
Data Science team in the National Statistician's Office (NSO) Analysis Unit (DASCH_RAPID@ons.gov.uk)
