#!/usr/bin/env Rscript
# Render the dashboard without RStudio
# Usage: Rscript render.R
#   - Renders dashboard.Rmd to HTML
#   - Automatically opens in your default browser

cat(">> Rendering Team Tesla Dashboard...\n")

rmarkdown::render(
  input = "dashboard.Rmd",
  output_file = "dashboard.html",
  knit_root_dir = getwd(),
  quiet = FALSE
)

cat(">> Done! Opening in browser...\n")
browseURL("dashboard.html")
