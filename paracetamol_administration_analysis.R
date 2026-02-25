####Script Details####
## Title:        Paracetamol Administration analysis
## Description:  Analyse Paracetamol administrations, derive 24h totals and <6h repeat dosing; includes paracetamol-equivalent dosing for paracetamol-containing products (2025 list) and route extraction from clinical display.
## Author:       Leon O'Hagan
## Github link:  https://github.com/leoninformatics/Paracetamol-Administration-Analysis
## Email:        lohagan@rotunda.ie
## Date:         2026-02-25

# ==============================
# 0) CONFIG
# ==============================
suppressPackageStartupMessages({
  library(tidyverse)
  library(readxl)
  library(lubridate)
  library(janitor)
  library(stringr)
  library(slider)
  library(writexl)
})

file_path <- "N:/Pharmacy/MN CMS/MN-CMS Pharmacy Reports/Paracetamol Analysis/Paracetamol_Admins_2025_12.xlsx"
sheet <- 1

threshold_mg <- 4000
min_interval_hours <- 6
tz <- "Europe/Dublin"

# Column mapping (use names as they appear in the source Excel header row)
column_map <- list(
  patient_id = "MRN",
  administration_datetime = "BEG_DT_TM",
  dose = "ADMIN_DOSE",
  dose_unit = "ADMIN_UNIT",
  medication_name = "MEDICATION_NAME",
  ward_location = "WARD",
  prescriber_order_id = "ORDER_ID",
  clinical_display_line = "CLINICAL_DISPLAY_LINE",
  hospital = "HOSPITAL",
  building = "BUILDING",
  result_status = "CE_RESULT_STATUS_DISP"
)

#list of paracetamol-containing products prescribed in 2025 (authoritative)
paracetamol_products_2025 <- c(
  "Calpol",
  "Calpol Fastmelts",
  "Co-codamol 30mg/500mg effervescent tablets",
  "Co-codamol 30mg/500mg tablets",
  "Codipar 15mg/500mg capsules",
  "Codipar 15mg/500mg effervescent tablets",
  "Excedrin",
  "Ixprim 37.5mg / 325mg effervescent tablets",
  "Ixprim 37.5mg / 325mg tablets",
  "Paracetamol",
  "Paracetamol (ANES)",
  "Paracetamol (ANES) 1 g",
  "Paracetamol 500mg/ Codeine 8mg/ Caffeine 30mg",
  "Paralief",
  "Solpadeine",
  "Solpadol 30mg/500mg caplets",
  "Solpadol 30mg/500mg effervescent tablets",
  "Tramadol 37.5mg / Paracetamol 325mg effervescent tablets",
  "Tylex 30mg/500mg capsules"
)

# Optional: add extra aliases (free-text brand fragments) for broader detection
brand_aliases <- c("calpol", "solpadeine", "solpadol", "tylex", "panadol", "tylenol", "paralief", "excedrin", "ixprim", "co-codamol", "codipar")

# ==============================
# 1)REQUIRED PACKAGES (check)
# ==============================
required_pkgs <- c("dplyr", "readxl", "lubridate", "stringr", "slider", "writexl", "tibble")
missing <- required_pkgs[!vapply(required_pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing) > 0) {
  stop(
    "Missing packages: ", paste(missing, collapse = ", "),
    "\nPlease install them (e.g., install.packages(...)) and re-run.",
    call. = FALSE
  )
}

# ==============================
# 2) HELPERS
# ==============================
coalesce_to_char <- function(x) {
  if (is.factor(x)) x <- as.character(x)
  if (inherits(x, "POSIXt") || inherits(x, "Date")) return(as.character(x))
  as.character(x)
}

# Case/whitespace-insensitive column resolver.
# Lets you map "MRN" even if the real column is "mrn" or " MRN " etc.
resolve_col <- function(df, wanted) {
  wanted_norm <- tolower(trimws(wanted))
  nms <- names(df)
  nms_norm <- tolower(trimws(nms))
  idx <- match(wanted_norm, nms_norm)
  if (!is.na(idx)) return(nms[idx])
  NA_character_
}

# Robust datetime parser (handles Excel numeric datetimes, and common strings)
parse_admin_datetime <- function(x, tz = "Europe/Dublin") {
  if (inherits(x, "POSIXt")) return(with_tz(x, tz))
  if (inherits(x, "Date")) return(as.POSIXct(x, tz = tz))
  
  # Excel numeric datetime
  if (is.numeric(x)) {
    # Excel origin (Windows): 1899-12-30
    return(as.POSIXct(x * 86400, origin = "1899-12-30", tz = tz))
  }
  
  x_chr <- str_squish(as.character(x))
  x_chr[x_chr == ""] <- NA_character_
  
  suppressWarnings(parse_date_time(
    x_chr,
    orders = c(
      "Ymd HMS", "Ymd HM", "Ymd",
      "dmY HMS", "dmY HM", "dmY",
      "dmy HMS", "dmy HM", "dmy",
      "d/b/Y HMS", "d/b/Y HM", "d/b/Y",
      "d-b-Y HMS", "d-b-Y HM", "d-b-Y",
      "d/m/y HMS", "d/m/y HM", "d/m/y",
      "d-b-y HMS", "d-b-y HM", "d-b-y",
      "d/b/y HMS", "d/b/y HM", "d/b/y",
      "d-b-y H:M:S", "d-b-Y H:M:S",
      "d-b-y H:M", "d-b-Y H:M"
    ),
    tz = tz
  ))
}

# Extract route from CLINICAL_DISPLAY_LINE:
# "DOSE: ... - ROUTE: intraVENOUS - infusion - ..."  -> "intravenous"
extract_route_from_display <- function(x) {
  x_chr <- as.character(x)
  route <- str_match(x_chr, "(?i)\\broute\\s*:\\s*([^\\-]+)")[, 2]
  route <- str_squish(route)
  
  # normalise common routes
  route_lower <- tolower(route)
  route[route_lower %in% c("intravenous", "iv", "i.v.", "intra venous", "intra-venous")] <- "intravenous"
  route[route_lower %in% c("oral", "po", "p.o.")] <- "oral"
  
  route
}

# Flag if any of the supplied product strings appears in medication_name (fixed match, case-insensitive via tolower)
product_used_flag <- function(med_name, patterns_lower) {
  med <- tolower(trimws(as.character(med_name)))
  ifelse(is.na(med), FALSE, Reduce(`|`, lapply(patterns_lower, function(p) str_detect(med, fixed(p)))))
}

# Capture first matched product name (best-effort)
product_used_name <- function(med_name, patterns_lower) {
  med <- tolower(trimws(as.character(med_name)))
  out <- rep(NA_character_, length(med))
  for (p in patterns_lower) {
    hit <- !is.na(med) & is.na(out) & str_detect(med, fixed(p))
    out[hit] <- p
  }
  str_to_title(out)
}

# Parse paracetamol strength from medication_name if present.
# Returns numeric mg of PARACETAMOL per unit (tablet/capsule/etc.) when detectable; else NA.
parse_paracetamol_strength_mg <- function(med_name) {
  med <- tolower(as.character(med_name))
  
  # 1) explicit "paracetamol XXX mg"
  parac_mg <- suppressWarnings(as.numeric(str_match(med, "(?i)paracetamol\\s*([0-9]+\\.?[0-9]*)\\s*mg")[, 2]))
  
  # 2) explicit "paracetamol ... X g"
  parac_g <- suppressWarnings(as.numeric(str_match(med, "(?i)paracetamol[^0-9]*([0-9]+\\.?[0-9]*)\\s*g\\b")[, 2]))
  parac_g_mg <- ifelse(!is.na(parac_g), parac_g * 1000, NA_real_)
  
  # 3) fallback: take the max of all mg numbers in the string (common combos: 30mg/500mg)
  mg_nums <- str_extract_all(med, "([0-9]+\\.?[0-9]*)\\s*mg")
  fallback <- vapply(mg_nums, function(v) {
    if (length(v) == 0) return(NA_real_)
    vals <- suppressWarnings(as.numeric(str_match(v, "([0-9]+\\.?[0-9]*)")[, 2]))
    if (all(is.na(vals))) return(NA_real_)
    max(vals, na.rm = TRUE)
  }, numeric(1))
  fallback[is.infinite(fallback)] <- NA_real_
  
  out <- parac_mg
  out[is.na(out)] <- parac_g_mg[is.na(out)]
  out[is.na(out)] <- fallback[is.na(out)]
  out
}

# Vectorised paracetamol-equivalent mg:
# - grams -> mg
# - if dose is a count and strength can be parsed, dose * strength_mg
# - Solpadeine capsules override: 1 capsule = 500 mg (your rule)
paracetamol_equiv_mg <- function(med_name, dose, unit = NA_character_) {
  med <- tolower(trimws(as.character(med_name)))
  unit <- tolower(trimws(as.character(unit)))
  dose_num <- suppressWarnings(as.numeric(dose))
  
  out <- dose_num
  
  # grams -> mg
  is_g <- !is.na(unit) & unit %in% c("g", "gram", "grams")
  out[is_g] <- dose_num[is_g] * 1000
  
  # Strength-based conversion when unit indicates a count of items
  strength_mg <- parse_paracetamol_strength_mg(med_name)
  is_count_unit <- !is.na(unit) & str_detect(unit, "tab|tablet|cap|capsule|caplet|effervescent|fastmelt|supp|sachet|dose|unit|spray|lozenge")
  use_strength <- !is.na(strength_mg) & !is.na(dose_num) & is_count_unit
  out[use_strength] <- dose_num[use_strength] * strength_mg[use_strength]
  
  # Solpadeine capsules override (requested): 500 mg per capsule
  is_solv_cap <- str_detect(med, "solpadeine") & !is.na(unit) & str_detect(unit, "cap")
  out[is_solv_cap] <- dose_num[is_solv_cap] * 500
  
  out
}

# ==============================
# 3) IMPORT
# ==============================
message("[1/7] Importing Excel...")
if (!file.exists(file_path)) {
  stop("Input file does not exist. Update 'file_path' in the config section and rerun.", call. = FALSE)
}

raw <- read_excel(file_path, sheet = sheet, col_names = TRUE, .name_repair = "minimal")
names(raw) <- trimws(names(raw))

message("[2/7] Resolving mapped columns and building working dataset...")

# Resolve mapped names against actual dataframe names (case/trim-insensitive)
resolved <- lapply(column_map, function(nm) if (is.na(nm)) NA_character_ else resolve_col(raw, nm))

# Quick sanity check (required columns)
required_cols <- c("patient_id", "administration_datetime", "dose")
missing_req <- names(resolved)[is.na(resolved[names(resolved) %in% required_cols])]
if (length(missing_req) > 0) {
  stop(
    "Could not resolve required columns: ", paste(missing_req, collapse = ", "),
    "\nCheck the header row / skip parameter / column_map values.",
    call. = FALSE
  )
}

working <- raw %>%
  transmute(
    patient_id = .data[[resolved$patient_id]],
    administration_datetime_raw = .data[[resolved$administration_datetime]],
    dose_raw = .data[[resolved$dose]],
    dose_unit = if (!is.na(resolved$dose_unit)) .data[[resolved$dose_unit]] else NA_character_,
    medication_name = if (!is.na(resolved$medication_name)) .data[[resolved$medication_name]] else NA_character_,
    ward_location = if (!is.na(resolved$ward_location)) .data[[resolved$ward_location]] else NA_character_,
    prescriber_order_id = if (!is.na(resolved$prescriber_order_id)) .data[[resolved$prescriber_order_id]] else NA_character_,
    clinical_display_line = if (!is.na(resolved$clinical_display_line)) .data[[resolved$clinical_display_line]] else NA_character_,
    hospital = if (!is.na(resolved$hospital)) .data[[resolved$hospital]] else NA_character_,
    building = if (!is.na(resolved$building)) .data[[resolved$building]] else NA_character_,
    result_status = if (!is.na(resolved$result_status)) .data[[resolved$result_status]] else NA_character_
  )

# ==============================
# 3b) CLEAN BASICS
# ==============================
message("[2.5/7] Cleaning datetimes and basic fields...")

n_before <- nrow(working)

working <- working %>%
  mutate(
    patient_id = coalesce_to_char(patient_id),
    medication_name = coalesce_to_char(medication_name),
    dose_unit = coalesce_to_char(dose_unit),
    ward_location = coalesce_to_char(ward_location),
    prescriber_order_id = coalesce_to_char(prescriber_order_id),
    clinical_display_line = coalesce_to_char(clinical_display_line),
    administration_datetime = parse_admin_datetime(administration_datetime_raw, tz = tz),
    dose_raw = coalesce_to_char(dose_raw)
  )

bad_dt <- sum(is.na(working$administration_datetime))
message("Rows with unparseable administration datetime: ", bad_dt, " / ", n_before)

# ==============================
# 4) FILTER + ROUTE + BRAND/PRODUCT + PARACETAMOL-EQUIV DOSE
# ==============================
message("[3/7] Filtering paracetamol-containing products (2025 list), extracting route, product flags, and deriving paracetamol-equivalent dose...")

# Normalised patterns for product detection
product_patterns_lower <- unique(tolower(trimws(paracetamol_products_2025)))
alias_patterns_lower <- unique(tolower(trimws(brand_aliases)))

# Generic synonym safety-net
generic_paracetamol_pattern <- regex("paracetamol|acetaminophen|aceta[min]?ophen|\\bapap\\b", ignore_case = TRUE)

# Filter if medication_name is available (and not all NA)
if (!all(is.na(working$medication_name)) && any(nzchar(na.omit(working$medication_name)))) {
  before_med_filter <- nrow(working)
  working <- working %>%
    filter(
      str_detect(tolower(coalesce_to_char(medication_name)), generic_paracetamol_pattern) |
        product_used_flag(medication_name, product_patterns_lower) |
        product_used_flag(medication_name, alias_patterns_lower)
    )
  after_med_filter <- nrow(working)
  message("Rows kept after medication filter: ", after_med_filter, " / ", before_med_filter)
} else {
  message("No medication_name mapped/present; assuming source already contains only paracetamol-containing administrations.")
}

working <- working %>%
  mutate(
    # Route extracted from clinical display line
    route_extracted = if (!all(is.na(clinical_display_line))) extract_route_from_display(clinical_display_line) else NA_character_,
    
    # Product/brand flags based on your 2025 list (+ aliases)
    brand_used_flag = if (!all(is.na(medication_name))) {
      product_used_flag(medication_name, product_patterns_lower) | product_used_flag(medication_name, alias_patterns_lower)
    } else FALSE,
    
    brand_used_name = if (!all(is.na(medication_name))) {
      # Prefer an exact product match; otherwise fall back to alias capture
      nm <- product_used_name(medication_name, product_patterns_lower)
      nm2 <- product_used_name(medication_name, alias_patterns_lower)
      ifelse(!is.na(nm), nm, nm2)
    } else NA_character_,
    
    # Paracetamol-equivalent dose in mg (strength parsing + Solpadeine capsule override)
    dose_mg = paracetamol_equiv_mg(medication_name, dose_raw, dose_unit)
  )

bad_dose <- sum(is.na(working$dose_mg))
message("Rows with unparseable dose (dose_mg is NA): ", bad_dose, " / ", nrow(working))

# ==============================
# 5) DERIVED METRICS: <6h + rolling 24h total
# ==============================
message("[4/7] Calculating time-since-previous and rolling 24h totals...")

working <- working %>%
  filter(!is.na(patient_id) & nzchar(patient_id)) %>%
  filter(!is.na(administration_datetime)) %>%
  arrange(patient_id, administration_datetime) %>%
  group_by(patient_id) %>%
  mutate(
    time_since_prev_hours = as.numeric(difftime(administration_datetime, lag(administration_datetime), units = "hours")),
    within_6h_flag = !is.na(time_since_prev_hours) & time_since_prev_hours < min_interval_hours
  ) %>%
  ungroup()

# Rolling 24h total (inclusive of current administration)
working <- working %>%
  group_by(patient_id) %>%
  arrange(administration_datetime, .by_group = TRUE) %>%
  mutate(
    rolling_24h_window_start = administration_datetime - hours(24),
    rolling_24h_total_mg = slider::slide_index_dbl(
      .x = dose_mg,
      .i = administration_datetime,
      .f = ~ sum(.x, na.rm = TRUE),
      .before = hours(24),
      .complete = FALSE
    ),
    exceeds_24h_threshold_flag = !is.na(rolling_24h_total_mg) & rolling_24h_total_mg > threshold_mg
  ) %>%
  ungroup()

# ==============================
# 6) SUMMARIES + EXCEPTIONS
# ==============================
message("[5/7] Building summary tables and exceptions...")

clean_admins <- working %>%
  arrange(patient_id, administration_datetime)

patient_day_summary <- clean_admins %>%
  mutate(calendar_date = as.Date(administration_datetime, tz = tz)) %>%
  group_by(patient_id, calendar_date) %>%
  summarise(
    total_dose_mg_calendar_day = sum(dose_mg, na.rm = TRUE),
    max_rolling_24h_total_mg_that_day = suppressWarnings(max(rolling_24h_total_mg, na.rm = TRUE)),
    n_admins = dplyr::n(),
    n_within_6h = sum(within_6h_flag, na.rm = TRUE),
    any_exceeds_threshold = any(exceeds_24h_threshold_flag, na.rm = TRUE),
    any_brand_used = any(brand_used_flag, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    max_rolling_24h_total_mg_that_day = ifelse(is.infinite(max_rolling_24h_total_mg_that_day), NA_real_, max_rolling_24h_total_mg_that_day)
  )

exceptions <- clean_admins %>%
  filter(within_6h_flag | exceeds_24h_threshold_flag) %>%
  mutate(
    exception_reason = case_when(
      within_6h_flag & exceeds_24h_threshold_flag ~ paste0("<", min_interval_hours, "h since previous AND 24h total > ", threshold_mg, "mg"),
      within_6h_flag ~ paste0("<", min_interval_hours, "h since previous dose"),
      exceeds_24h_threshold_flag ~ paste0("Rolling 24h total > ", threshold_mg, "mg"),
      TRUE ~ NA_character_
    )
  )

# Data quality sheet
data_quality <- tibble::tibble(
  metric = c(
    "Rows imported",
    "Rows after medication filter (if applied)",
    "Rows with unparseable datetime",
    "Rows with unparseable dose_mg",
    "Datetime min",
    "Datetime max",
    "Threshold mg (rolling 24h)",
    "Min interval hours (< flag)"
  ),
  value = c(
    n_before,
    nrow(clean_admins),
    bad_dt,
    bad_dose,
    as.character(suppressWarnings(min(clean_admins$administration_datetime, na.rm = TRUE))),
    as.character(suppressWarnings(max(clean_admins$administration_datetime, na.rm = TRUE))),
    threshold_mg,
    min_interval_hours
  )
)

# ==============================
# 7) EXPORT
# ==============================
message("[6/7] Exporting Excel workbook...")

out_path <- file.path(getwd(), paste0("Paracetamol_Admin_Analysis_", format(Sys.Date(), "%Y%m%d"), ".xlsx"))

writexl::write_xlsx(
  x = list(
    clean_admins = clean_admins,
    patient_day_summary = patient_day_summary,
    exceptions = exceptions,
    data_quality = data_quality
  ),
  path = out_path
)

message("[7/7] Done. Output written to: ", out_path)

# ==============================
# (Optional) QUICK PLOTS
# ==============================
#If you want quick visuals, uncomment:
if (requireNamespace("ggplot2", quietly = TRUE)) {
  library(ggplot2)
  ggplot(clean_admins %>% filter(!is.na(time_since_prev_hours)),
         aes(x = time_since_prev_hours)) +
    geom_histogram(bins = 50) +
    labs(title = "Time since previous Paracetamol administration (hours)",
         x = "Hours", y = "Count")

  top_patients <- clean_admins %>%
    group_by(patient_id) %>%
    summarise(max_roll = max(rolling_24h_total_mg, na.rm = TRUE), .groups = "drop") %>%
    arrange(desc(max_roll)) %>%
    slice_head(n = 5) %>%
    pull(patient_id)

  ggplot(clean_admins %>% filter(patient_id %in% top_patients),
         aes(x = administration_datetime, y = rolling_24h_total_mg)) +
    geom_line() +
    facet_wrap(~ patient_id, scales = "free_x") +
    labs(title = "Rolling 24h Paracetamol total (top 5 patients)",
         x = "Administration time", y = "Rolling 24h total (mg)")
}
