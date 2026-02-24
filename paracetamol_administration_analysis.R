####Script Details####
## Title:       Paracetamol Administration analysis
## Description: Production-quality script for cleaning and analysing Paracetamol administration records from Excel exports.
## Author:      
## Github link:      
## Email:      lohagan@rotunda.ie
## Date:        2026-02-24

suppressPackageStartupMessages({
  library(tidyverse)
  library(readxl)
  library(lubridate)
  library(janitor)
  library(stringr)
  library(slider)
  library(writexl)
})

# ==============================
# 0) USER INPUTS / CONFIG
# ==============================
file_path <- "path/to/paracetamol_administrations.xlsx"
sheet <- 1
threshold_mg <- 4000
min_interval_hours <- 6
tz <- "Europe/Dublin"

# Column mapping (set values to the source column names as they appear in Excel)
column_map <- list(
  patient_id = "patient_id",
  administration_datetime = "administration_datetime",
  dose_mg = "dose_mg",
  medication_name = NA_character_,   # optional
  route = NA_character_,             # optional
  ward_location = NA_character_,     # optional
  prescriber_order_id = NA_character_ # optional
)

# Output config
output_dir <- "."
output_prefix <- paste0("paracetamol_analysis_", format(Sys.Date(), "%Y%m%d"))

# ==============================
# 1) PACKAGE CHECKS (NO AUTO-INSTALL)
# ==============================
required_pkgs <- c("tidyverse", "readxl", "lubridate", "janitor", "stringr", "slider", "writexl")
missing_pkgs <- required_pkgs[!vapply(required_pkgs, requireNamespace, FUN.VALUE = logical(1), quietly = TRUE)]
if (length(missing_pkgs) > 0) {
  stop(
    paste0(
      "Missing required packages: ", paste(missing_pkgs, collapse = ", "),
      "\nPlease install them in your R environment before running this script.",
      "\nExample: install.packages(c('", paste(missing_pkgs, collapse = "','"), "'))"
    ),
    call. = FALSE
  )
}

# ==============================
# 2) HELPER FUNCTIONS
# ==============================
validate_column_mapping <- function(data, column_map) {
  required_fields <- c("patient_id", "administration_datetime", "dose_mg")
  missing_fields <- required_fields[!required_fields %in% names(column_map)]
  if (length(missing_fields) > 0) {
    stop("column_map is missing required mapping entries: ", paste(missing_fields, collapse = ", "), call. = FALSE)
  }

  required_columns <- unlist(column_map[required_fields], use.names = FALSE)
  absent_columns <- required_columns[!required_columns %in% names(data)]
  if (length(absent_columns) > 0) {
    stop(
      "The following mapped required columns were not found in the input sheet: ",
      paste(absent_columns, collapse = ", "),
      call. = FALSE
    )
  }

  optional_keys <- setdiff(names(column_map), required_fields)
  invalid_optional <- optional_keys[
    !is.na(unlist(column_map[optional_keys])) &
      !(unlist(column_map[optional_keys]) %in% names(data))
  ]
  if (length(invalid_optional) > 0) {
    warning(
      "Some optional mapped columns were not found and will be ignored: ",
      paste(invalid_optional, collapse = ", "),
      call. = FALSE
    )
  }
}

coalesce_to_char <- function(x) {
  if (is.null(x)) return(rep(NA_character_, 0))
  as.character(x)
}

parse_admin_datetime <- function(x, tz = "Europe/Dublin") {
  if (inherits(x, "POSIXct")) {
    return(with_tz(x, tzone = tz))
  }

  if (inherits(x, "Date")) {
    return(as.POSIXct(x, tz = tz))
  }

  if (is.numeric(x)) {
    # Excel serial dates (Windows origin)
    return(as.POSIXct(x * 86400, origin = "1899-12-30", tz = tz))
  }

  x_chr <- trimws(as.character(x))
  x_chr[x_chr %in% c("", "NA", "N/A", "NULL", "null", "na")] <- NA_character_

  parsed <- suppressWarnings(
    parse_date_time(
      x_chr,
      orders = c(
        "Ymd HMS", "Ymd HM", "Ymd",
        "dmY HMS", "dmY HM", "dmY",
        "mdY HMS", "mdY HM", "mdY",
        "dmy HMS", "dmy HM", "dmy",
        "mdy HMS", "mdy HM", "mdy",
        "ymd HMS", "ymd HM", "ymd",
        "d-b-Y HMS", "d-b-Y HM", "d-b-Y",
        "d/b/Y HMS", "d/b/Y HM", "d/b/Y",
        "Y-m-d H:M:S", "Y-m-d H:M", "Y-m-d",
        "d.m.Y H:M:S", "d.m.Y H:M", "d.m.Y"
      ),
      tz = tz,
      exact = FALSE,
      quiet = TRUE
    )
  )

  as.POSIXct(parsed, tz = tz)
}

parse_dose_to_mg <- function(x) {
  if (is.numeric(x)) return(as.numeric(x))

  x_chr <- tolower(trimws(as.character(x)))
  x_chr[x_chr %in% c("", "na", "n/a", "null")] <- NA_character_

  number <- suppressWarnings(readr::parse_number(x_chr, locale = readr::locale(decimal_mark = ".")))
  unit <- case_when(
    str_detect(x_chr, "\\bmcg\\b|\\bug\\b|microgram") ~ "mcg",
    str_detect(x_chr, "\\bg\\b|gram") ~ "g",
    str_detect(x_chr, "\\bmg\\b|milligram") ~ "mg",
    TRUE ~ "mg" # default assumption
  )

  dose_mg <- case_when(
    is.na(number) ~ NA_real_,
    unit == "g" ~ number * 1000,
    unit == "mcg" ~ number / 1000,
    TRUE ~ number
  )

  as.numeric(dose_mg)
}

build_exception_reason <- function(within_6h_flag, exceeds_24h_threshold_flag) {
  case_when(
    isTRUE(within_6h_flag) & isTRUE(exceeds_24h_threshold_flag) ~ "<6h since previous; exceeds 24h threshold",
    isTRUE(within_6h_flag) ~ "<6h since previous",
    isTRUE(exceeds_24h_threshold_flag) ~ "Exceeds rolling 24h threshold",
    TRUE ~ NA_character_
  )
}

# ==============================
# 3) IMPORT
# ==============================
if (!file.exists(file_path)) {
  stop("Input file does not exist. Update 'file_path' in the config section and rerun.", call. = FALSE)
}

message("[1/7] Reading Excel data...")
raw <- read_excel(path = file_path, sheet = sheet) %>%
  clean_names()

if (nrow(raw) == 0) stop("Input data has 0 rows; nothing to process.", call. = FALSE)

# normalize column_map values to janitor::clean_names() format
column_map <- purrr::map_chr(column_map, ~ {
  if (is.na(.x)) return(NA_character_)
  janitor::make_clean_names(.x)
})

validate_column_mapping(raw, column_map)

message("[2/7] Applying column mapping and initial selection...")
working <- raw %>%
  transmute(
    patient_id = .data[[column_map$patient_id]],
    administration_datetime_raw = .data[[column_map$administration_datetime]],
    dose_raw = .data[[column_map$dose_mg]],
    medication_name = if (!is.na(column_map$medication_name) && column_map$medication_name %in% names(raw)) .data[[column_map$medication_name]] else NA,
    route = if (!is.na(column_map$route) && column_map$route %in% names(raw)) .data[[column_map$route]] else NA,
    ward_location = if (!is.na(column_map$ward_location) && column_map$ward_location %in% names(raw)) .data[[column_map$ward_location]] else NA,
    prescriber_order_id = if (!is.na(column_map$prescriber_order_id) && column_map$prescriber_order_id %in% names(raw)) .data[[column_map$prescriber_order_id]] else NA
  )

initial_n <- nrow(working)

# ==============================
# 4) FILTER TO PARACETAMOL (if medication column provided)
# ==============================
message("[3/7] Filtering to Paracetamol/Acetaminophen records (when medication_name is available)...")
paracetamol_pattern <- regex("paracetamol|acetaminophen|aceta[min]?ophen|\\bapap\\b|calpol|panadol|tylenol", ignore_case = TRUE)

if (!all(is.na(working$medication_name))) {
  before_med_filter <- nrow(working)
  working <- working %>% filter(str_detect(coalesce_to_char(medication_name), paracetamol_pattern))
  after_med_filter <- nrow(working)
  message("Rows kept after medication filter: ", after_med_filter, " / ", before_med_filter)
} else {
  message("No medication_name mapped/present; assuming source already contains only Paracetamol administrations.")
}

# ==============================
# 5) CLEAN DATETIME / DOSE + DATA QUALITY FILTERS
# ==============================
message("[4/7] Parsing administration datetime and dose...")
working <- working %>%
  mutate(
    patient_id = as.character(patient_id),
    administration_datetime = parse_admin_datetime(administration_datetime_raw, tz = tz),
    dose_mg = parse_dose_to_mg(dose_raw)
  )

missing_patient_id_n <- sum(is.na(working$patient_id) | trimws(working$patient_id) == "")
invalid_datetime_n <- sum(is.na(working$administration_datetime))
invalid_dose_n <- sum(is.na(working$dose_mg) | !is.finite(working$dose_mg) | working$dose_mg <= 0)

clean_admins <- working %>%
  filter(!(is.na(patient_id) | trimws(patient_id) == "")) %>%
  filter(!is.na(administration_datetime)) %>%
  filter(!is.na(dose_mg), is.finite(dose_mg), dose_mg > 0) %>%
  arrange(patient_id, administration_datetime)

clean_n <- nrow(clean_admins)
dropped_total_n <- initial_n - clean_n

# ==============================
# 6) DERIVE FLAGS + ROLLING CALCULATIONS
# ==============================
message("[5/7] Computing intervals, rolling 24h totals, and exception flags...")
clean_admins <- clean_admins %>%
  group_by(patient_id) %>%
  arrange(administration_datetime, .by_group = TRUE) %>%
  mutate(
    time_since_prev_hours = as.numeric(difftime(administration_datetime, lag(administration_datetime), units = "hours")),
    within_6h_flag = !is.na(time_since_prev_hours) & time_since_prev_hours < min_interval_hours,
    rolling_24h_total_mg = slide_index_dbl(
      .x = dose_mg,
      .i = administration_datetime,
      .f = ~ sum(.x, na.rm = TRUE),
      .before = hours(24),
      .after = 0,
      .complete = FALSE
    ),
    exceeds_24h_threshold_flag = rolling_24h_total_mg > threshold_mg,
    rolling_24h_window_start = administration_datetime - hours(24)
  ) %>%
  ungroup() %>%
  select(
    patient_id,
    administration_datetime,
    dose_mg,
    time_since_prev_hours,
    within_6h_flag,
    rolling_24h_total_mg,
    exceeds_24h_threshold_flag,
    rolling_24h_window_start,
    medication_name,
    route,
    ward_location,
    prescriber_order_id
  )

# Optional calendar-day summary
message("[6/7] Building patient-day summary and exceptions...")
patient_day_summary <- clean_admins %>%
  mutate(calendar_date = as.Date(administration_datetime, tz = tz)) %>%
  group_by(patient_id, calendar_date) %>%
  summarise(
    total_dose_mg_calendar_day = sum(dose_mg, na.rm = TRUE),
    max_rolling_24h_total_mg_that_day = max(rolling_24h_total_mg, na.rm = TRUE),
    n_admins = n(),
    n_within_6h = sum(within_6h_flag, na.rm = TRUE),
    any_exceeds_threshold = any(exceeds_24h_threshold_flag, na.rm = TRUE),
    .groups = "drop"
  )

exceptions <- clean_admins %>%
  filter(within_6h_flag | exceeds_24h_threshold_flag) %>%
  mutate(exception_reason = purrr::map2_chr(within_6h_flag, exceeds_24h_threshold_flag, build_exception_reason))

# Data quality sheet
date_min <- if (clean_n > 0) min(clean_admins$administration_datetime, na.rm = TRUE) else as.POSIXct(NA)
date_max <- if (clean_n > 0) max(clean_admins$administration_datetime, na.rm = TRUE) else as.POSIXct(NA)

data_quality <- tibble(
  metric = c(
    "input_rows_initial",
    "rows_after_cleaning",
    "rows_dropped_total",
    "rows_missing_or_blank_patient_id",
    "rows_invalid_or_missing_datetime",
    "rows_invalid_or_missing_dose",
    "exceptions_rows",
    "distinct_patients",
    "analysis_timezone",
    "analysis_threshold_mg",
    "analysis_min_interval_hours",
    "datetime_range_start",
    "datetime_range_end"
  ),
  value = c(
    as.character(initial_n),
    as.character(clean_n),
    as.character(dropped_total_n),
    as.character(missing_patient_id_n),
    as.character(invalid_datetime_n),
    as.character(invalid_dose_n),
    as.character(nrow(exceptions)),
    as.character(dplyr::n_distinct(clean_admins$patient_id)),
    tz,
    as.character(threshold_mg),
    as.character(min_interval_hours),
    as.character(date_min),
    as.character(date_max)
  )
)

# ==============================
# 7) EXPORT + OPTIONAL PLOTS
# ==============================
message("[7/7] Exporting outputs to Excel and creating optional plots...")
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

output_xlsx <- file.path(output_dir, paste0(output_prefix, ".xlsx"))

writexl::write_xlsx(
  x = list(
    clean_admins = clean_admins,
    patient_day_summary = patient_day_summary,
    exceptions = exceptions,
    data_quality = data_quality
  ),
  path = output_xlsx
)

# Optional visual outputs
plot1_path <- file.path(output_dir, paste0(output_prefix, "_time_since_prev_hist.png"))
plot2_path <- file.path(output_dir, paste0(output_prefix, "_rolling24h_top5.png"))

if (nrow(clean_admins %>% filter(!is.na(time_since_prev_hours))) > 0) {
  p1 <- clean_admins %>%
    filter(!is.na(time_since_prev_hours)) %>%
    ggplot(aes(x = time_since_prev_hours)) +
    geom_histogram(binwidth = 1, fill = "steelblue", color = "white") +
    geom_vline(xintercept = min_interval_hours, linetype = "dashed", color = "red") +
    labs(
      title = "Distribution of time since previous administration",
      x = "Hours since previous administration",
      y = "Count"
    ) +
    theme_minimal()

  ggsave(filename = plot1_path, plot = p1, width = 10, height = 6, dpi = 300)
}

top5_patients <- clean_admins %>%
  group_by(patient_id) %>%
  summarise(max_rolling = max(rolling_24h_total_mg, na.rm = TRUE), .groups = "drop") %>%
  slice_max(order_by = max_rolling, n = 5, with_ties = FALSE) %>%
  pull(patient_id)

if (length(top5_patients) > 0) {
  p2 <- clean_admins %>%
    filter(patient_id %in% top5_patients) %>%
    ggplot(aes(x = administration_datetime, y = rolling_24h_total_mg, color = patient_id, group = patient_id)) +
    geom_line(alpha = 0.8) +
    geom_point(size = 1.2) +
    geom_hline(yintercept = threshold_mg, linetype = "dashed", color = "red") +
    facet_wrap(~ patient_id, scales = "free_x") +
    labs(
      title = "Rolling 24-hour total dose (Top 5 patients by max rolling total)",
      x = "Administration datetime",
      y = "Rolling 24h total (mg)",
      color = "Patient"
    ) +
    theme_minimal() +
    theme(legend.position = "none")

  ggsave(filename = plot2_path, plot = p2, width = 12, height = 8, dpi = 300)
}

message("Analysis complete.")
message("Excel output: ", output_xlsx)
if (file.exists(plot1_path)) message("Plot saved: ", plot1_path)
if (file.exists(plot2_path)) message("Plot saved: ", plot2_path)
