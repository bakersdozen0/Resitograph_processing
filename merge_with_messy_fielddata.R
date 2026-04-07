# AUTOMATED METADATA MERGING SUITE
# Version: Standardized "BLOCK" Anchor Pipeline
# Description: Merges detrended Resistograph data (.xlsx) with field assessment sheets 
# using "BLOCK" anchors to map non-unique scan IDs to their corresponding sub-directories.

# SECTION 1: SETUP & LIBRARIES ####
library(tidyxl)
library(openxlsx)
library(writexl)
library(readxl)
library(here)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)

# SECTION 2: CORE MERGING ALGORITHM ####
process_block_anchored_experiment <- function(layout_path, processed_path, experiment_name) {
  
  message(paste("\n======================================================="))
  message(paste("--- Processing:", experiment_name, "---"))
  message(paste("======================================================="))
  
  # ---------------------------------------------------------
  # PART A: READ FIELD LAYOUT (Bounding Box + Vertical Lanes)
  # ---------------------------------------------------------
  cells <- xlsx_cells(layout_path, sheets = 1)
  
  # 1. Find all "Plot" or "Tree" columns globally 
  global_plot_cols <- cells %>%
    filter(grepl("^(Plot|Tree)", str_trim(character), ignore.case = TRUE)) %>%
    pull(col) %>%
    unique()
  
  if(length(global_plot_cols) == 0) stop("CRITICAL: Could not find any 'Plot' or 'Tree' headers in the sheet.")
  
  # 2. Locate every distinct Block instance
  block_anchors <- cells %>%
    filter(grepl("^BLOCK\\s*\\d+", character, ignore.case = TRUE)) %>%
    arrange(row, col)
  
  if(nrow(block_anchors) == 0) stop("CRITICAL: No 'BLOCK' anchors found.")
  
  unique_block_cols <- sort(unique(block_anchors$col))
  field_data_list <- list()
  
  # 3. Iterate through each Block's Bounding Box
  for (i in seq_len(nrow(block_anchors))) {
    
    curr_block <- block_anchors[i, ]
    b_row <- curr_block$row
    b_col <- curr_block$col
    
    # Extract Folder Name (Cell immediately right)
    folder_cell <- cells %>% filter(row == b_row, col == b_col + 1)
    folder_val <- if(nrow(folder_cell) > 0) coalesce(folder_cell$character[1], as.character(folder_cell$numeric[1])) else "UNKNOWN"
    
    # Calculate Box Boundaries using Vertical Lanes
    next_lanes <- unique_block_cols[unique_block_cols > b_col]
    b_right <- if(length(next_lanes) > 0) next_lanes[1] - 1 else max(cells$col)
    
    blocks_below <- block_anchors %>% filter(row > b_row, col >= b_col, col <= b_right) %>% arrange(row)
    b_bottom <- if(nrow(blocks_below) > 0) blocks_below$row[1] - 1 else max(cells$row)
    
    # Grab all cells in this box
    block_cells <- cells %>%
      filter(row > b_row, row <= b_bottom, col >= b_col, col <= b_right)
    
    # Extract data using generalized offsets (0=Plot, 1=RESI, 2=Cull)
    for (p_col in global_plot_cols) {
      
      extracted_data <- block_cells %>%
        filter(col >= p_col, col <= p_col + 2) %>% # Adjusted offset limit
        mutate(
          val = coalesce(as.character(numeric), character),
          offset = as.character(col - p_col)
        ) %>%
        filter(!is.na(val)) %>%
        select(row, offset, val)
      
      if(nrow(extracted_data) == 0) next
      
      wide_data <- extracted_data %>%
        pivot_wider(names_from = offset, values_from = val)
      
      # Ensure all offset columns (0 through 2) exist
      for (off in 0:2) { 
        if (!as.character(off) %in% names(wide_data)) {
          wide_data[[as.character(off)]] <- NA_character_
        }
      }
      
      if (!"0" %in% names(wide_data) || all(is.na(wide_data$`0`))) next
      
      # Clean up and apply generalized names
      clean_data <- wide_data %>%
        rename(plot_label = `0`, RESI_ID = `1`, Cull = `2`) %>% 
        # Added the $ so it only drops the exact header, keeping "Plot 12" safe!
        filter(!grepl("^(Plot|Tree)$", str_trim(plot_label), ignore.case = TRUE)) %>% 
        filter(!is.na(plot_label)) %>%
        mutate(folder_name = folder_val)
      
      field_data_list[[length(field_data_list) + 1]] <- clean_data
    }
  }
  
  # 4. Combine into one tidy dataframe
  raw_field_data <- bind_rows(field_data_list) %>%
    mutate(
      Tree_ID = str_extract(plot_label, "\\d+"),
      folder_match_key = toupper(gsub("[^A-Z0-9]", "", folder_name)),
      RESI_ID_clean = suppressWarnings(as.character(as.numeric(RESI_ID)))
    ) %>%
    # DROP FILLERS: Instantly remove anything that didn't have a number in the Plot column
    filter(!is.na(Tree_ID)) %>%
    distinct(folder_match_key, Tree_ID, .keep_all = TRUE)
  
  # ---------------------------------------------------------
  # DIAGNOSTICS & AUDITS (Updated for D, M, Dead, Missing)
  # ---------------------------------------------------------
  regex_dead <- "^\\s*(Dead|D)\\s*$"
  regex_missing <- "^\\s*(Missing|M)\\s*$"
  regex_swept <- "^\\s*Swept\\s*$"
  regex_all_skips <- "^\\s*(Dead|Missing|D|M|Swept)\\s*$"
  
  count_dead <- sum(grepl(regex_dead, raw_field_data$RESI_ID, ignore.case = TRUE), na.rm = TRUE)
  count_missing <- sum(grepl(regex_missing, raw_field_data$RESI_ID, ignore.case = TRUE), na.rm = TRUE)
  count_swept <- sum(grepl(regex_swept, raw_field_data$RESI_ID, ignore.case = TRUE), na.rm = TRUE)
  count_numeric <- sum(!is.na(raw_field_data$RESI_ID_clean))
  total_found <- count_dead + count_missing + count_swept + count_numeric
  
  message(paste("  -> DIAGNOSTIC: Found", count_dead, "'Dead',", count_missing, "'Missing', and", count_swept, "'Swept' trees."))
  message(paste("  -> DIAGNOSTIC: Found", count_numeric, "trees with standard numeric RESI scans."))
  message(paste("  -> DIAGNOSTIC: Total categorized trees extracted:", total_found))
  
  # AUDIT 1: Unrecognized / Typo Checks
  unrecognized_scans <- raw_field_data %>%
    filter(
      is.na(RESI_ID_clean),
      !grepl(regex_all_skips, RESI_ID, ignore.case = TRUE),
      !is.na(RESI_ID)
    )
  
  if(nrow(unrecognized_scans) > 0) {
    message("\n  [!] AUDIT: Found plots with unrecognized text in the RESI column (Typo?):")
    print(unrecognized_scans %>% select(plot_label, folder_name, RESI_ID) %>% as.data.frame())
  }
  
  # Remove non-scans prior to merge
  field_data_filtered <- raw_field_data %>%
    filter(!is.na(RESI_ID_clean))
  
  # ---------------------------------------------------------
  # PART B: READ PROCESSED RESISTOGRAPH DATA
  # ---------------------------------------------------------
  sheet_names <- excel_sheets(processed_path)
  
  tidy_data <- sheet_names %>%
    set_names() %>% 
    map_df(~read_excel(path = processed_path, sheet = .x, col_types = "text"), .id = "source_sheet") %>%
    mutate(
      Sample_date_num = suppressWarnings(as.numeric(Sample_date)),
      Sample_date = if_else(
        !is.na(Sample_date_num),
        format(excel_numeric_to_date(Sample_date_num, include_time = TRUE), "%d/%m/%Y %H:%M"),
        Sample_date
      ),
      file_number_extract = str_extract(File_name, "\\d+(?=\\.rgp$)"),
      file_number_extract = if_else(is.na(file_number_extract), str_extract(File_name, "\\d+"), file_number_extract),
      RESI_ID_clean = as.character(as.numeric(file_number_extract)),
      folder_match_key = toupper(gsub("[^A-Z0-9]", "", source_sheet))
    ) %>%
    select(-any_of(c("Tree_ID", "Plot")))
  
  # ---------------------------------------------------------
  # AUDIT 2: Failed Matches
  # ---------------------------------------------------------
  failed_matches <- field_data_filtered %>%
    anti_join(tidy_data, by = c("folder_match_key", "RESI_ID_clean"))
  
  if(nrow(failed_matches) > 0) {
    message(paste("\n  [!] AUDIT:", nrow(failed_matches), "numeric field scans failed to match with the machine data:"))
    print(failed_matches %>% select(plot_label, folder_name, RESI_ID) %>% as.data.frame())
  }
  
  # ---------------------------------------------------------
  # PART C: MERGE (Inner Join)
  # ---------------------------------------------------------
  final_merged <- field_data_filtered %>%
    inner_join(tidy_data, by = c("folder_match_key", "RESI_ID_clean")) %>%
    select(
      Tree_ID, plot_label, folder_name, RESI_ID, any_of(c("Cull", "Field_Date", "Assessor", "Device")), 
      everything(),
      -folder_match_key, -file_number_extract, -source_sheet, -Sample_date_num, -RESI_ID_clean
    ) %>%
    mutate(across(everything(), as.character)) %>%
    arrange(as.numeric(Tree_ID))
  
  return(final_merged)
}

# ==============================================================================
# SECTION 3: MAIN EXECUTION BLOCK 
# ==============================================================================
# Instructions: Update the paths below and run this block to process your data.

# 1. Define Details
experiment_name     <- "Moray 62" # Change this as needed
processed_data_file <- here("Processed data", paste0(experiment_name, "_Automated_Output.xlsx"))
layout_map_file     <- here(experiment_name, paste0(experiment_name, "_Resi_fieldsheet.xlsx"))
output_file_name    <- paste0(experiment_name, "_Combined_Final.xlsx")

# 2. Execute Pipeline
tryCatch({
  
  df_final <- process_block_anchored_experiment(
    layout_path = layout_map_file, 
    processed_path = processed_data_file, 
    experiment_name = experiment_name
  )
  
  # 3. Save to Excel
  if (nrow(df_final) > 0) {
    write_xlsx(df_final, here("Processed Data", output_file_name))
    message(paste("\nSUCCESS: Saved", output_file_name, "with", nrow(df_final), "records."))
  } else {
    message("\nWARNING: Process finished but resulted in 0 records. Check folder name matching.")
  }
  
}, error = function(e) {
  message("\nAn error occurred during processing: ", e$message)
})
