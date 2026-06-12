# =============================================================================
# Author:      Sylvester Kotak Vinther
# Date:        15-12-2025
# Version:     5.5 (Updated with unweighted option and option to fit offset 'e')

# Description:
# This code fits single/double exponential decay and optionally the integral. Then calculates the RLU at "custom_x_value" (after x-shift).
# Options added for unweighted fitting and adding a constant offset 'e' to the exponential curves.
# The code outputs the fitted regressions (+ integral + integral regressions if enabled). RLU at custom_x_value, max read value (from the TECAN excel)
# The high_val is set to the greatest value between max_read and custom_x_value. The chosen source value is highlighted with light blue.
# Fit quality is annotated and marked with a smooth color gradient from red-yellow-green where yellow corresponds to an R^2 of ~0.9.
# The x-axis (time) can be shifted such 0 = the actual V_t0 of the experiment.
# REMEMBER: If the filter is used all blue cells in the input .xlsx file should be multiplied by 1.12
# =============================================================================
rm(list = ls())
# First time only: Install the following packages with: install.packages(c("readxl", "ggplot2", "minpack.lm", "foreach", "doParallel", "openxlsx", "tools"))
library(readxl)
library(ggplot2)
library(minpack.lm)
library(foreach)
library(doParallel)
library(openxlsx)
library(tools)

start_time <- Sys.time()

# =============================================================================
# SETTINGS & INPUT FILES
# =============================================================================
fit_type <- "double"            # Options: "double" or "single"
include_offset <- FALSE         # TRUE adds a constant factor '+ e' to the fit, FALSE keeps it pure exponential
calculate_integral <- FALSE      # When FALSE, the integral is neither calculated nor plotted.
output_failed_fits <- TRUE      # TRUE: Output rows with data but failed fits; FALSE: Skip them
weighting_method <- "1/y^2"     # Options: "1/y^2", "1/y", or "unweighted"
x_shift <- 1.2                  # Shifts the x-values such x_shift is equal to t0. This is done to keep V_t0 at time=0.  For injection at 180µL/s this equals 1.2s shift.
custom_x_value <- 0             # The x-value (on the SHIFTED time axis) for calculating Y. This should be set to the time where RLU_max is expected (on the SHIFTED time axis) 
input_file <- "C:/Users/thebf_000/Desktop/Videnskabelig assistent/Dilutions Problem/Getting NaCl concentration correct (it is 100mM).xlsx"
#input file Path. Output will be set to same location. NOTE Windows uses "\" for file paths. R uses "/"
sheet_names <- excel_sheets(input_file) 
r2_threshold_for_t0_highlight <- 0.95 

# =============================================================================
# METADATA SETTINGS
# =============================================================================
# Enable this to read names from a specified column in the input Excel file.
use_metadata <- TRUE

# Specify the column containing the sample name. You can use a column letter (e.g., "B") or number (e.g., 2).
# This name will be included in the output file for easy identification.
name_col <- "B"   # Column containing the sample name or condition.

# =============================================================================
# SET UP PARALLEL PROCESSING
# =============================================================================
n_cores <- parallel::detectCores() - 1
if (n_cores < 1) n_cores <- 1 
cl <- makeCluster(n_cores)
registerDoParallel(cl)
cat("Using", n_cores, "cores for parallel processing.\n")

# =============================================================================
# PROCESS SHEETS (PARALLEL)
# =============================================================================
sheet_results <- foreach(sheet = sheet_names,
                         .packages = c("readxl", "ggplot2", "minpack.lm")
) %dopar% {
  
  sheet_coefficients_list <- list()
  sheet_y_values_list <- list()
  sheet_plots <- list()
  sheet_integrated_plots <- list()
  
  # ----- Read Sheet Data -----
  cat("Attempting to read sheet:", sheet, "\n")
  excel_data_full <- tryCatch({
    read_excel(input_file, sheet = sheet, col_names = FALSE, .name_repair = "minimal")
  }, error = function(e) {
    cat("Error reading sheet:", sheet, "-", e$message, "\n")
    return(NULL)
  })
  
  if (is.null(excel_data_full) || nrow(excel_data_full) == 0 || ncol(excel_data_full) == 0) {
    cat("Skipping sheet", sheet, "- sheet is empty or has insufficient data.\n")
    return(NULL)
  }
  
  # Prepare for metadata reading
  name_col_num <- NA_integer_
  if (exists("use_metadata") && use_metadata) {
    .convert_col_to_numeric <- function(col_ref) {
      if (is.null(col_ref) || is.na(col_ref)) return(NA_integer_)
      if (is.numeric(col_ref)) return(as.integer(col_ref))
      if (is.character(col_ref)) {
        val <- 0; s <- toupper(trimws(col_ref))
        for (i in 1:nchar(s)) {
          char <- substr(s, i, i)
          if (!char %in% LETTERS) return(NA_integer_)
          val <- val * 26 + (as.integer(charToRaw(char)) - 64)
        }
        return(val)
      }
      return(NA_integer_)
    }
    name_col_num <- .convert_col_to_numeric(name_col)
    
    if (!is.na(name_col_num) && name_col_num > ncol(excel_data_full)) {
      cat("Warning: Name column", name_col, "is out of bounds for sheet", sheet, ". Metadata will not be read.\n")
      name_col_num <- NA_integer_
    }
  }
  
  first_col_char <- tryCatch({
    as.character(excel_data_full[[1]])
  }, error = function(e) {
    cat("Warning: Could not convert column 1 to character for sheet", sheet, ". It might be empty or non-standard.\n")
    return(character(0)) 
  })
  
  if (length(first_col_char) == 0 && nrow(excel_data_full) > 0) { 
    cat("Skipping sheet", sheet, "- Column 1 could not be processed to find 'Time [s]'.\n")
    return(NULL)
  }
  
  time_s_matches <- which(trimws(first_col_char) == "Time [s]")
  
  if (length(time_s_matches) == 0) {
    cat("Skipping sheet", sheet, "- 'Time [s]' not found in column A.\n")
    return(NULL)
  }
  if (length(time_s_matches) > 1) {
    cat("Warning: Multiple 'Time [s]' found in column A for sheet", sheet, ". Using the first one at row", time_s_matches[1], ".\n")
  }
  time_row_index_full <- time_s_matches[1]
  
  data_start_col_index <- 4 
  num_y_data_rows <- 96
  
  if (time_row_index_full > nrow(excel_data_full) || ncol(excel_data_full) < data_start_col_index) {
    cat("Skipping sheet", sheet, "- 'Time [s]' row or data columns are out of bounds.\n")
    return(NULL)
  }
  
  x_values_raw_from_sheet <- excel_data_full[time_row_index_full, data_start_col_index:ncol(excel_data_full)]
  x_values_numeric_raw <- as.numeric(unlist(x_values_raw_from_sheet))
  
  non_na_x_indices_in_slice_master <- which(!is.na(x_values_numeric_raw))
  if (length(non_na_x_indices_in_slice_master) == 0) {
    cat("Skipping sheet", sheet, "- No valid numeric X-values found in 'Time[s]' row from column D onwards.\n")
    return(NULL)
  }
  
  x_master_unfiltered_from_row <- x_values_numeric_raw[non_na_x_indices_in_slice_master]
  master_x_axis_min_shifted <- min(x_master_unfiltered_from_row, na.rm = TRUE) - x_shift
  master_x_axis_max_shifted <- max(x_master_unfiltered_from_row, na.rm = TRUE) - x_shift
  
  y_data_columns_in_full_excel <- (data_start_col_index - 1) + non_na_x_indices_in_slice_master
  y_data_block_start_row_full <- time_row_index_full + 1
  actual_y_data_block_end_row_full <- min(time_row_index_full + num_y_data_rows, nrow(excel_data_full))
  
  if (y_data_block_start_row_full > actual_y_data_block_end_row_full) {
    cat("Skipping sheet", sheet, "- No valid Y data rows available based on calculated range.\n")
    return(NULL)
  }
  
  y_data_block_raw <- excel_data_full[y_data_block_start_row_full:actual_y_data_block_end_row_full,
                                      y_data_columns_in_full_excel,
                                      drop = FALSE]
  y_data_block_numeric <- as.data.frame(lapply(y_data_block_raw, function(col_val) as.numeric(as.character(col_val))))
  
  dataset_row_indices_in_block <- which(rowSums(!is.na(y_data_block_numeric)) > 0)
  if (length(dataset_row_indices_in_block) == 0) {
    cat("No data rows with non-NA Y-values found in the y-data block for sheet", sheet, ".\n")
    return(NULL)
  }
  
  for (i_loop_idx in seq_along(dataset_row_indices_in_block)) {
    current_block_row_idx <- dataset_row_indices_in_block[i_loop_idx]
    actual_excel_row <- (y_data_block_start_row_full - 1) + current_block_row_idx
    
    # Read name metadata for the current data row
    current_name <- NA_character_
    if (exists("use_metadata") && use_metadata && !is.na(name_col_num)) {
      current_name <- tryCatch(as.character(excel_data_full[[name_col_num]][actual_excel_row]), error = function(e) NA_character_)
    }
    
    y_values_this_dataset_raw_slice <- as.numeric(y_data_block_numeric[current_block_row_idx, ])
    
    # --- MODIFIED SECTION START: Handle 'OVER' values at the start ---
    # We find the index of the LAST valid data point. This effectively defines the length of the assay for this row.
    # NAs before this index (e.g. "OVER" at the start) are treated as skipped time points.
    
    valid_data_indices <- which(!is.na(y_values_this_dataset_raw_slice))
    
    if (length(valid_data_indices) == 0) {
      next
    }
    
    last_valid_index <- max(valid_data_indices)
    
    # We generate a time sequence that spans from Start to End, but with a number of steps equal to the position of the last valid point.
    # e.g., if the last valid point is at index 15, we generate 15 time steps.
    # If index 1 was "OVER", we simply don't use the 1st time step, we use the 2nd one for the 2nd data point.
    full_time_seq_for_row <- seq(from = master_x_axis_min_shifted, 
                                 to = master_x_axis_max_shifted, 
                                 length.out = last_valid_index)
    
    # Filter: Keep only the X values where we actually have Y data
    x_values_for_fitting_scaled <- full_time_seq_for_row[valid_data_indices]
    y_values_no_na <- y_values_this_dataset_raw_slice[valid_data_indices]
    
    # --- MODIFIED SECTION END ---
    
    current_max_read <- max(y_values_no_na, na.rm = TRUE)
    data_no_na <- data.frame(x = x_values_for_fitting_scaled, y = y_values_no_na)
    
    # --- UPDATED WEIGHTING LOGIC ---
    if (weighting_method == "1/y") {
      calculated_weights <- 1 / abs(data_no_na$y)
    } else if (weighting_method == "1/y^2") {
      calculated_weights <- 1 / (data_no_na$y^2)
    } else {
      calculated_weights <- rep(1, nrow(data_no_na)) # unweighted
    }
    
    # Safety check: if y is 0 (unlikely in RLU but possible), weight becomes Inf. 
    calculated_weights[is.infinite(calculated_weights)] <- 0
    
    if (fit_type == "double") {
      a_candidates <- c(10000, 100000, 500000, 1000000, 50000000, 1000000000); b_candidates <- c(-1, -0.5, -0.3, -0.1, -0.05)
      c_candidates <- c(10000, 100000, 500000, 50000000, 1000000000); d_candidates <- c(-0.25, -0.03, -0.035)
      if (include_offset) {
        e_candidates <- c(0, 100000, 10000000)
        param_grid <- expand.grid(a = a_candidates, b = b_candidates, c = c_candidates, d = d_candidates, e = e_candidates)
        lower_bounds <- c(1e+3, -1, 1e+3, -1, -1e6)
        upper_bounds <- c(2e+11, -0.001, 2e+11, -0.001, 1e8)
      } else {
        param_grid <- expand.grid(a = a_candidates, b = b_candidates, c = c_candidates, d = d_candidates)
        lower_bounds <- c(1e+2, -1, 1e+2, -1)
        upper_bounds <- c(2e+11, -0.001, 2e+11, -0.001)
      }
    } else {
      a_candidates <- c(10000, 100000, 500000,10000000); b_candidates <- c(-1, -0.8, -0.5, -0.3, -0.1, -0.08, -0.05, -0.03, -0.01)
      if (include_offset) {
        e_candidates <- c(0, 100, 1000)
        param_grid <- expand.grid(a = a_candidates, b = b_candidates, e = e_candidates)
        lower_bounds <- c(1e+1, -10, -1e6)
        upper_bounds <- c(2e+11, -1e-3, 1e8)
      } else {
        param_grid <- expand.grid(a = a_candidates, b = b_candidates)
        lower_bounds <- c(1e+1, -10)
        upper_bounds <- c(2e+11, -1e-3)
      }
    }
    
    best_model <- NULL; best_rss <- Inf
    for (j in 1:nrow(param_grid)) {
      start_params <- as.list(param_grid[j,])
      tryCatch({
        if (fit_type == "double") {
          if (include_offset) {
            model_formula <- {y ~ a * exp(b * x) + c * exp(d * x) + e} 
            current_start_params <- list(a=start_params$a, b=start_params$b, c=start_params$c, d=start_params$d, e=start_params$e) 
          } else {
            model_formula <- {y ~ a * exp(b * x) + c * exp(d * x)} 
            current_start_params <- list(a=start_params$a, b=start_params$b, c=start_params$c, d=start_params$d) 
          }
        } else {
          if (include_offset) {
            model_formula <- {y ~ a * exp(b * x) + e}
            current_start_params <- list(a=start_params$a, b=start_params$b, e=start_params$e) 
          } else {
            model_formula <- {y ~ a * exp(b * x)}
            current_start_params <- list(a=start_params$a, b=start_params$b) 
          }
        }
        
        model <- nlsLM(model_formula, data = data_no_na, start = current_start_params,
                       lower = lower_bounds, upper = upper_bounds, weights = calculated_weights, control = nls.control(maxiter = 500))
        rss <- sum(residuals(model)^2)
        if (rss < best_rss) { best_rss <- rss; best_model <- model }
      }, error = function(e) { })
    }
    
    internal_y_col_name <- if (custom_x_value != 0) "y_at_custom_x_val" else "y_at_x_0"
    internal_v_col_name <- "V_at_custom_x_val"
    
    if (is.null(best_model)) {
      if (output_failed_fits) {
        coeff_data_failed <- list(
          Sheet = sheet, Excel_Row = actual_excel_row, Name = current_name,
          a = NA_real_, `-k1` = NA_real_, c = NA_real_, `-k2` = NA_real_, e = NA_real_,
          Equation = "Fit Failed", Goodness_of_fit = NA_real_, Halftime = NA_real_,
          Integrated_Equation = NA_character_, Integral = NA_character_
        )
        coeff_data_failed[[internal_v_col_name]] <- NA_real_
        
        y_row_failed_content <- list(Sheet = sheet, Excel_Row = actual_excel_row, Max_read = NA_real_)
        y_row_failed_content[[internal_y_col_name]] <- NA_real_
        
        sheet_coefficients_list[[length(sheet_coefficients_list) + 1]] <- as.data.frame(coeff_data_failed, check.names = FALSE)
        sheet_y_values_list[[length(sheet_y_values_list) + 1]] <- as.data.frame(y_row_failed_content, check.names = FALSE)
      }
      next
    }
    
    r_squared_val <- 1 - (sum(residuals(best_model)^2) / sum((data_no_na$y - mean(data_no_na$y))^2))
    
    if (!is.finite(r_squared_val)) {
      if (output_failed_fits) {
        coeff_data_failed <- list(
          Sheet = sheet, Excel_Row = actual_excel_row, Name = current_name,
          a = NA_real_, `-k1` = NA_real_, c = NA_real_, `-k2` = NA_real_, e = NA_real_,
          Equation = "Fit Failed (Invalid R^2)", Goodness_of_fit = NA_real_, Halftime = NA_real_,
          Integrated_Equation = NA_character_, Integral = NA_character_
        )
        coeff_data_failed[[internal_v_col_name]] <- NA_real_
        
        y_row_failed_content <- list(Sheet = sheet, Excel_Row = actual_excel_row, Max_read = NA_real_)
        y_row_failed_content[[internal_y_col_name]] <- NA_real_
        
        sheet_coefficients_list[[length(sheet_coefficients_list) + 1]] <- as.data.frame(coeff_data_failed, check.names = FALSE)
        sheet_y_values_list[[length(sheet_y_values_list) + 1]] <- as.data.frame(y_row_failed_content, check.names = FALSE)
      }
      next 
    }
    
    coeffs <- coef(best_model)
    param_a <- as.numeric(coeffs["a"]) 
    param_b <- as.numeric(coeffs["b"]) 
    param_c <- if (fit_type == "double" && "c" %in% names(coeffs)) as.numeric(coeffs["c"]) else NA_real_
    param_d <- if (fit_type == "double" && "d" %in% names(coeffs)) as.numeric(coeffs["d"]) else NA_real_
    param_e <- if (include_offset && "e" %in% names(coeffs)) as.numeric(coeffs["e"]) else NA_real_
    val_e   <- if (is.na(param_e)) 0 else param_e
    
    # --- SWAP LOGIC: Ensure param_d is ALWAYS the less variable (slower) rate constant ---
    if (fit_type == "double" && !is.na(param_c) && !is.na(param_d)) {
      if (param_b > param_d) {
        temp_a <- param_a; param_a <- param_c; param_c <- temp_a
        temp_b <- param_b; param_b <- param_d; param_d <- temp_b
      }
    }
    # ------------------------------------------------------------------------------------
    
    V_at_custom_x <- NA_real_
    if (fit_type == "double") {
      if (!is.na(param_a) && !is.na(param_b) && !is.na(param_c) && !is.na(param_d)) {
        V_at_custom_x <- param_a * param_b * exp(param_b * custom_x_value) + param_c * param_d * exp(param_d * custom_x_value)
      }
    } else {
      if (!is.na(param_a) && !is.na(param_b)) {
        V_at_custom_x <- param_a * param_b * exp(param_b * custom_x_value)
      }
    }
    
    y_at_shifted_x_is_0 <- NA_real_
    if (!is.na(param_a) && !is.na(param_b)) {
      term1_at_0 <- param_a * exp(param_b * 0)
      term2_at_0 <- if (fit_type == "double" && !is.na(param_c) && !is.na(param_d)) param_c * exp(param_d * 0) else 0 
      y_at_shifted_x_is_0 <- term1_at_0 + term2_at_0 + val_e
    }
    
    value_for_output_y_column <- NA_real_
    if (custom_x_value != 0) {
      target_x_for_pred <- custom_x_value
      if (fit_type == "double") {
        if (!is.na(param_a) && !is.na(param_b) && !is.na(param_c) && !is.na(param_d)) { 
          value_for_output_y_column <- param_a * exp(param_b * target_x_for_pred) + param_c * exp(param_d * target_x_for_pred) + val_e
        }
      } else { 
        if (!is.na(param_a) && !is.na(param_b)) {
          value_for_output_y_column <- param_a * exp(param_b * target_x_for_pred) + val_e
        }
      }
    } else {
      value_for_output_y_column <- y_at_shifted_x_is_0
    }
    
    equation_term1 <- paste(round(param_a, 10), "*exp(", round(param_b, 10), "*x)") 
    equation <- if (fit_type == "double" && !is.na(param_c) && !is.na(param_d)) { 
      paste("y =", equation_term1, "+", round(param_c, 10), "*exp(", round(param_d, 10), "*x)")
    } else {
      paste("y =", equation_term1)
    }
    if (include_offset && !is.na(param_e)) {
      if (param_e >= 0) {
        equation <- paste(equation, "+", round(param_e, 10))
      } else {
        equation <- paste(equation, "-", abs(round(param_e, 10)))
      }
    }
    
    integrated_equation <- NA_character_
    if (calculate_integral) {
      integ_term_str <- function(coeff, rate) {
        if (is.na(coeff) || is.na(rate)) return("NA_TERM")
        paste0("(", round(coeff,10), "/", round(-rate,10), ")(1-exp(", round(rate,10), "x))")
      }
      term1_integ_eq_str <- integ_term_str(param_a, param_b) 
      
      if (fit_type == "double") {
        if (term1_integ_eq_str == "NA_TERM") {
          integrated_equation <- "NA"
        } else {
          term2_integ_eq_str <- integ_term_str(param_c, param_d) 
          if (term2_integ_eq_str == "NA_TERM") integrated_equation <- term1_integ_eq_str
          else integrated_equation <- paste0(term1_integ_eq_str, "+", term2_integ_eq_str)
        }
      } else { 
        integrated_equation <- if(term1_integ_eq_str == "NA_TERM") "NA" else term1_integ_eq_str
      }
      if (include_offset && !is.na(param_e) && integrated_equation != "NA") {
        if (param_e >= 0) {
          integrated_equation <- paste0(integrated_equation, "+", round(param_e, 10), "x")
        } else {
          integrated_equation <- paste0(integrated_equation, "-", abs(round(param_e, 10)), "x")
        }
      }
    }
    
    halftime <- NA_real_ 
    if (!is.na(y_at_shifted_x_is_0) && y_at_shifted_x_is_0 > 0) {
      half_val <- y_at_shifted_x_is_0 / 2
      halftime_upper_bound <- max(1e6, master_x_axis_max_shifted + 100, na.rm=TRUE)
      
      halftime_func <- if (fit_type == "double" && !is.na(param_c) && !is.na(param_d) && !is.na(param_a) && !is.na(param_b)) { 
        function(x_val) param_a * exp(param_b * x_val) + param_c * exp(param_d * x_val) + val_e - half_val
      } else if (!is.na(param_a) && !is.na(param_b)) { 
        function(x_val) param_a * exp(param_b * x_val) + val_e - half_val
      } else { NULL }
      
      if (!is.null(halftime_func)) {
        tryCatch({
          f_lower <- halftime_func(0)
          f_upper <- halftime_func(halftime_upper_bound)
          if (!is.na(f_lower) && !is.na(f_upper) && sign(f_lower) != sign(f_upper)) {
            halftime_result <- uniroot(halftime_func, lower = 0, upper = halftime_upper_bound)$root
            if (!is.na(halftime_result)) halftime <- halftime_result
          }
        }, error = function(e) { })
      }
    }
    
    integral_val <- NA_character_
    if (calculate_integral) {
      if (include_offset && !is.na(param_e) && abs(param_e) > 1e-9) {
        # The integral to infinity of a non-zero offset diverges
        integral_val <- ifelse(param_e > 0, "Inf", "-Inf")
      } else {
        calc_term_integral_inf <- function(coeff_val, rate_val) {
          if (is.na(coeff_val) || is.na(rate_val)) return(NA_real_)
          if (abs(rate_val) < 1e-9) return(if (coeff_val == 0) 0 else Inf * sign(coeff_val))
          return(-coeff_val / rate_val)
        }
        
        integral_val_calc <- tryCatch({
          val_to_return <- NA_real_
          term1_inf <- calc_term_integral_inf(param_a, param_b) 
          
          if (fit_type == "double") {
            if (!is.na(term1_inf)) {
              term2_inf <- calc_term_integral_inf(param_c, param_d) 
              if (!is.na(term2_inf)) val_to_return <- term1_inf + term2_inf
              else val_to_return <- term1_inf 
            } 
          } else { 
            val_to_return <- term1_inf
          }
          val_to_return 
        }, error = function(e) { NA_real_ })
        
        if (!is.na(integral_val_calc) && is.finite(integral_val_calc)) {
          integral_val <- as.character(integral_val_calc)
        } else if (!is.na(integral_val_calc) && is.infinite(integral_val_calc)) {
          integral_val <- ifelse(integral_val_calc > 0, "Inf", "-Inf")
        }
      }
    }
    
    coeff_data_list_content <- list(
      Sheet = sheet, Excel_Row = actual_excel_row, Name = current_name,
      a = param_a, `-k1` = param_b, c = NA_real_, `-k2` = NA_real_, e = param_e,
      Equation = equation, Goodness_of_fit = r_squared_val, Halftime = halftime,
      Integrated_Equation = integrated_equation, Integral = integral_val
    )
    coeff_data_list_content[[internal_v_col_name]] <- V_at_custom_x
    
    if (fit_type == "double") {
      coeff_data_list_content$c <- param_c
      coeff_data_list_content$`-k2` <- param_d 
    }
    sheet_coefficients_list[[length(sheet_coefficients_list) + 1]] <- as.data.frame(coeff_data_list_content, check.names = FALSE)
    
    y_row_list_content <- list(Sheet = sheet, Excel_Row = actual_excel_row, Max_read = current_max_read)
    y_row_list_content[[internal_y_col_name]] <- value_for_output_y_column
    sheet_y_values_list[[length(sheet_y_values_list) + 1]] <- as.data.frame(y_row_list_content, check.names = FALSE)
    
    df_plot_points <- data_no_na
    x_seq_for_plot_line <- if (abs(master_x_axis_max_shifted - master_x_axis_min_shifted) < 1e-9) master_x_axis_min_shifted else seq(master_x_axis_min_shifted, master_x_axis_max_shifted, length.out = 100)
    df_plot_line <- data.frame(x = x_seq_for_plot_line, y = predict(best_model, newdata = data.frame(x = x_seq_for_plot_line)))
    plot_coord_lims <- NULL
    if (is.finite(master_x_axis_min_shifted) && is.finite(master_x_axis_max_shifted)) {
      plot_coord_lims <- if (abs(master_x_axis_max_shifted - master_x_axis_min_shifted) < 1e-9) {
        coord_cartesian(xlim = c(master_x_axis_min_shifted - 0.1, master_x_axis_max_shifted + 0.1))
      } else {
        coord_cartesian(xlim = c(master_x_axis_min_shifted, master_x_axis_max_shifted))
      }
    }
    p <- ggplot() + geom_point(data = df_plot_points, aes(x = x, y = y)) + geom_line(data = df_plot_line, aes(x = x, y = y), color = "red") +
      plot_coord_lims + labs(title = paste("Fit:", sheet, "Row", actual_excel_row, ",", current_name), x = paste0("Time (scaled & shifted by ", x_shift, ")"), y = "RLU") +
      annotate("text", x = Inf, y = Inf, hjust = 1.1, vjust = 2, label = paste(equation, "\nR² =", round(r_squared_val, 3)), size = 3) + theme_minimal()
    sheet_plots[[length(sheet_plots) + 1]] <- p
    if (calculate_integral) {
      safe_integral_term_plot <- function(coeff, rate, t_val) { 
        if (is.na(coeff) || is.na(rate)) return(rep(NA_real_, length(t_val)))
        if (abs(rate) < 1e-9) return(coeff * t_val)
        return((coeff / rate) * (exp(rate * t_val) - 1))
      }
      integrated_y_values_plot <- NA_real_ 
      term1_plot_vals <- safe_integral_term_plot(param_a, param_b, x_seq_for_plot_line) 
      if (fit_type == "double" && !is.na(param_c) && !is.na(param_d)) { 
        term2_plot_vals <- safe_integral_term_plot(param_c, param_d, x_seq_for_plot_line)
        if(anyNA(term1_plot_vals) || anyNA(term2_plot_vals)) integrated_y_values_plot <- rep(NA_real_, length(x_seq_for_plot_line))
        else integrated_y_values_plot <- term1_plot_vals + term2_plot_vals
      } else { integrated_y_values_plot <- term1_plot_vals }
      
      if (include_offset && !is.na(param_e) && !anyNA(integrated_y_values_plot)) {
        integrated_y_values_plot <- integrated_y_values_plot + (param_e * x_seq_for_plot_line)
      }
      
      df_integrated_plot <- data.frame(x = x_seq_for_plot_line, y = integrated_y_values_plot)
      p_integrated <- ggplot(df_integrated_plot, aes(x = x, y = y)) + geom_line(color = "blue") +
        plot_coord_lims + labs(title = paste("Integral:", sheet, "Row", actual_excel_row, ",", current_name), x = paste0("Time (scaled & shifted by ", x_shift, ")"), y = "Integrated RLU") +
        annotate("text", x = Inf, y = Inf, hjust = 1.1, vjust = 2, label = integrated_equation, size = 3) + theme_minimal()
      sheet_integrated_plots[[length(sheet_integrated_plots) + 1]] <- p_integrated
    }
  } 
  
  final_sheet_coeff <- if (length(sheet_coefficients_list) > 0) do.call(rbind, sheet_coefficients_list) else NULL
  final_sheet_yvals <- if (length(sheet_y_values_list) > 0) do.call(rbind, sheet_y_values_list) else NULL
  
  if (is.null(final_sheet_coeff) || nrow(final_sheet_coeff) == 0) return(NULL) 
  
  list(coeff = final_sheet_coeff, y_vals = final_sheet_yvals, plots = sheet_plots, integrated_plots = sheet_integrated_plots)
} 
# =============================================================================
# POST-PROCESSING: COMBINE RESULTS
# =============================================================================
stopCluster(cl)
cat("Parallel processing finished.\n")

sheet_results <- Filter(Negate(is.null), sheet_results) 

if (length(sheet_results) == 0) {
  stop("No valid results were generated from any sheets. Exiting.")
}

# --- Define Dynamic Column Names ---
output_y_col_name_for_df <- if (custom_x_value != 0) {
  paste0("cRLU_at_x", gsub("\\.", "_", as.character(custom_x_value))) 
} else {
  "cRLU_t0" 
}
output_v_col_name_for_df <- if (custom_x_value != 0) {
  paste0("V_at_x", gsub("\\.", "_", as.character(custom_x_value)))
} else {
  "V_t0"
}

internal_y_col_name_in_results <- if (custom_x_value != 0) "y_at_custom_x_val" else "y_at_x_0"
internal_v_col_name_in_results <- "V_at_custom_x_val" 

all_coefficients <- do.call(rbind, lapply(sheet_results, function(res) res$coeff))
all_y_values_raw <- do.call(rbind, lapply(sheet_results, function(res) res$y_vals)) 
all_plots <- unlist(lapply(sheet_results, function(res) res$plots), recursive = FALSE)
all_integrated_plots <- unlist(lapply(sheet_results, function(res) res$integrated_plots), recursive = FALSE)

if (is.null(all_coefficients) || nrow(all_coefficients) == 0 ) {
  stop("Coefficients data frame is empty even after processing all sheets. Exiting.")
}

final_output_df <- all_coefficients 

# --- Rename dynamic velocity column ---
if (internal_v_col_name_in_results %in% names(final_output_df)) {
  names(final_output_df)[names(final_output_df) == internal_v_col_name_in_results] <- output_v_col_name_for_df
}

if (!is.null(all_y_values_raw) && nrow(all_y_values_raw) > 0) {
  if (!internal_y_col_name_in_results %in% names(all_y_values_raw)) {
    all_y_values_raw[[internal_y_col_name_in_results]] <- NA_real_
  }
  if (!"Max_read" %in% names(all_y_values_raw)) {
    all_y_values_raw$Max_read <- NA_real_
  }
  y_data_to_merge <- all_y_values_raw[, c("Sheet", "Excel_Row", internal_y_col_name_in_results, "Max_read"), drop = FALSE]
  final_output_df <- merge(final_output_df, y_data_to_merge, 
                           by = c("Sheet", "Excel_Row"), all.x = TRUE, sort = FALSE)
  names(final_output_df)[names(final_output_df) == internal_y_col_name_in_results] <- output_y_col_name_for_df
} else {
  final_output_df[[output_y_col_name_for_df]] <- NA_real_
  final_output_df$Max_read <- NA_real_
}

if (!calculate_integral) {
  cols_to_remove <- c("Integrated_Equation", "Integral")
  final_output_df <- final_output_df[, !(names(final_output_df) %in% cols_to_remove), drop = FALSE]
}

# --- high_val calculation ---
if (!(output_y_col_name_for_df %in% names(final_output_df))) {
  final_output_df[[output_y_col_name_for_df]] <- NA_real_
}
if (!("Max_read" %in% names(final_output_df))) {
  final_output_df$Max_read <- NA_real_
}
if (!("Goodness_of_fit" %in% names(final_output_df))) {
  final_output_df$Goodness_of_fit <- NA_real_ 
}

final_output_df[[output_y_col_name_for_df]] <- as.numeric(as.character(final_output_df[[output_y_col_name_for_df]]))
final_output_df$Max_read <- as.numeric(as.character(final_output_df$Max_read))
final_output_df$Goodness_of_fit <- as.numeric(as.character(final_output_df$Goodness_of_fit))

Gof <- final_output_df$Goodness_of_fit
Y_col_values <- final_output_df[[output_y_col_name_for_df]]
Max_r_values <- final_output_df$Max_read

final_output_df$high_val <- ifelse(
  !is.na(Gof) & Gof >= r2_threshold_for_t0_highlight & !is.na(Y_col_values),
  pmax(Y_col_values, Max_r_values, na.rm = TRUE), 
  Max_r_values                                   
)
final_output_df$high_val <- ifelse(
  is.na(final_output_df$high_val) & !is.na(Y_col_values), 
  Y_col_values,
  final_output_df$high_val
)
final_output_df$high_val <- ifelse(
  is.na(final_output_df$high_val) & !is.na(Max_r_values),
  Max_r_values,
  final_output_df$high_val
)

# --- Strict Final Column Selection and Ordering ---
desired_cols <- c("Sheet", "Excel_Row")

# Add name column if enabled
if (exists("use_metadata") && use_metadata) {
  desired_cols <- c(desired_cols, "Name")
}

# Add core fit parameter and quality columns (added "e")
desired_cols <- c(desired_cols, "a", "-k1", "c", "-k2", "e", "Equation", "Goodness_of_fit", "Halftime")

# Add integral columns if enabled
if (calculate_integral) {
  desired_cols <- c(desired_cols, "Integrated_Equation", "Integral")
}

# Add calculated value columns
desired_cols <- c(desired_cols, output_v_col_name_for_df, output_y_col_name_for_df, "Max_read", "high_val")

# Ensure all desired columns exist, adding NA columns if missing.
for (col_name in desired_cols) {
  if (!(col_name %in% names(final_output_df))) {
    warning(paste("Desired column '", col_name, "' was not found. Adding it as NA.", sep=""))
    final_output_df[[col_name]] <- NA
  }
}

# Select and reorder the columns.
final_output_df <- final_output_df[, desired_cols, drop = FALSE]

# =============================================================================
# VERSIONED OUTPUT FILENAMES & SAVE
# =============================================================================
input_file_base <- file_path_sans_ext(basename(input_file)) 
output_dir <- dirname(input_file)
version <- 1; output_excel <- ""; output_pdf <- "" 
repeat {
  base_name <- paste0(input_file_base, "_output_v", version)
  output_excel <- file.path(output_dir, paste0(base_name, ".xlsx"))
  output_pdf <- file.path(output_dir, paste0(base_name, ".pdf"))
  if (!file.exists(output_excel) && !file.exists(output_pdf)) break
  version <- version + 1
}
cat("Output files will be version:", version, "\n")

wb <- createWorkbook()
addWorksheet(wb, "Output")
writeData(wb, "Output", final_output_df, keepNA = FALSE) 

headerStyle <- createStyle(textDecoration = "bold")
addStyle(wb, "Output", headerStyle, rows = 1, cols = 1:ncol(final_output_df))

# --- Apply SMOOTH GRADIENT coloring based on R^2 value ---

# Find column indices for R^2 and the target columns for coloring
r2_col_idx <- which(names(final_output_df) == "Goodness_of_fit")
v_col_idx <- which(names(final_output_df) == output_v_col_name_for_df)
crlu_col_idx <- which(names(final_output_df) == output_y_col_name_for_df)

# Check if all required columns exist before proceeding
if (length(r2_col_idx) > 0 && length(v_col_idx) > 0 && length(crlu_col_idx) > 0 && nrow(final_output_df) > 0) {
  
  # Define the colors for the gradient: Red -> Yellow -> Green
  gradient_colors <- c("#FFC7CE", "#FFEB9C", "#C6EFCE")
  color_palette_func <- colorRampPalette(gradient_colors)
  
  # Generate a fine-grained palette of 100 colors for the smooth gradient
  palette_100 <- color_palette_func(100)
  
  # Columns to apply the R^2 based coloring
  cols_to_color <- c(r2_col_idx, v_col_idx, crlu_col_idx)
  
  # Loop through each data row to apply conditional styling
  for (i in 1:nrow(final_output_df)) {
    r2_value <- final_output_df$Goodness_of_fit[i]
    excel_data_row <- i + 1 # +1 for the header row
    
    # Skip if R^2 is NA or not a finite number
    if (is.na(r2_value) || !is.finite(r2_value)) {
      next
    }
    
    # Determine Color based on Threshold logic
    if (r2_value < r2_threshold_for_t0_highlight) {
      # Strictly red below threshold (matches the start of the gradient)
      hex_color <- "#FFC7CE" 
    } else {
      # Gradient from Red (at threshold) to Green (at 1.0)
      # Clamp value to max 1.0
      val_for_grad <- min(r2_value, 1.0)
      
      # Calculate position in range[threshold, 1.0]
      range_span <- 1.0 - r2_threshold_for_t0_highlight
      if (range_span < 1e-9) range_span <- 1e-9 # Avoid division by zero if threshold is 1
      
      scale_factor <- (val_for_grad - r2_threshold_for_t0_highlight) / range_span
      color_index <- round(scale_factor * 99) + 1
      
      # Ensure index is within bounds (1-100)
      color_index <- max(1, min(100, color_index))
      hex_color <- palette_100[color_index]
    }
    
    # Create a new style on-the-fly for this specific color
    dynamic_style <- createStyle(fgFill = hex_color)
    
    # Apply the determined style to the R^2, V, and cRLU columns for the current row
    addStyle(wb, "Output", style = dynamic_style, rows = excel_data_row, cols = cols_to_color, gridExpand = TRUE)
  }
}

# --- Highlight the chosen source for 'high_val' ---
# Define a light blue style to mark the cell (cRLU or Max_read) that was used for the high_val column.
highlightStyle <- createStyle(fgFill = "#BDD7EE") # Light Blue

y_value_to_highlight_col_idx <- which(names(final_output_df) == output_y_col_name_for_df)
max_read_col_idx <- which(names(final_output_df) == "Max_read")

if (length(y_value_to_highlight_col_idx) > 0 && length(max_read_col_idx) > 0 && "high_val" %in% names(final_output_df) && nrow(final_output_df) > 0) {
  for (i in 1:nrow(final_output_df)) {
    excel_data_row <- i + 1 
    current_high_val <- final_output_df$high_val[i]
    if (!is.na(current_high_val)) {
      current_y_val_to_check <- final_output_df[[y_value_to_highlight_col_idx]][i]
      current_r2 <- if (length(r2_col_idx) > 0 && !is.na(final_output_df[[r2_col_idx]][i])) final_output_df[[r2_col_idx]][i] else NA_real_
      
      # Check if the calculated value was chosen
      if (!is.na(current_y_val_to_check) && abs(current_y_val_to_check - current_high_val) < 1e-9 && 
          !is.na(current_r2) && current_r2 >= r2_threshold_for_t0_highlight) {
        addStyle(wb, "Output", style = highlightStyle, rows = excel_data_row, cols = y_value_to_highlight_col_idx)
      } else { 
        # Otherwise, check if the max read value was chosen
        current_max_read <- final_output_df[[max_read_col_idx]][i]
        if (!is.na(current_max_read) && abs(current_max_read - current_high_val) < 1e-9) {
          addStyle(wb, "Output", style = highlightStyle, rows = excel_data_row, cols = max_read_col_idx)
        }
      }
    }
  }
}
saveWorkbook(wb, output_excel, overwrite = TRUE)
cat("Excel output saved to:", output_excel, "\n")

has_fit_plots <- length(all_plots) > 0 && any(!sapply(all_plots, is.null))
has_integral_plots <- calculate_integral && length(all_integrated_plots) > 0 && any(!sapply(all_integrated_plots, is.null))

if (has_fit_plots || has_integral_plots) {
  pdf(output_pdf, width=11, height=8.5)
  max_plots_count <- 0
  if(has_fit_plots) max_plots_count <- max(max_plots_count, length(all_plots))
  if(has_integral_plots) max_plots_count <- max(max_plots_count, length(all_integrated_plots))
  if(max_plots_count > 0) {
    for (i in 1:max_plots_count) {
      if (has_fit_plots && i <= length(all_plots) && !is.null(all_plots[[i]])) {
        tryCatch(print(all_plots[[i]]), error = function(e) cat("Error printing plot", i, ":", e$message, "\n"))
      }
      if (has_integral_plots && i <= length(all_integrated_plots) && !is.null(all_integrated_plots[[i]])) {
        tryCatch(print(all_integrated_plots[[i]]), error = function(e) cat("Error printing integrated plot", i, ":", e$message, "\n"))
      }
    }
  }
  dev.off()
  cat("Plots saved to:", output_pdf, "\n")
} else {
  cat("No plots were generated to save to PDF.\n")
}

end_time <- Sys.time()
cat("Total execution time:", format(end_time - start_time), "\n")