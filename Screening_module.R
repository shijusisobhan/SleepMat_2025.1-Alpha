

rm(list=ls())


packages = c("readxl", "dplyr", "openxlsx", "stringr")

## Load or install the required packges
package.check <- lapply(
  packages,
  FUN = function(x) {
    if (!require(x, character.only = TRUE)) {
      install.packages(x, dependencies = TRUE)
      library(x, character.only = TRUE)
    }
  }
)



# Browse and select the Excel file interactively
file_path <- file.choose()
mainDir <- dirname(file_path)

# Read the selected Excel file
file_info <- read_excel(file_path)


# Read CSV with genotype list (col1) and folders (col2)
# Read the CSV containing genotype and folder paths
# file_info <- read.csv("C:/Users/shijusis/OneDrive - Michigan Medicine/Desktop/Screening_module/M18_info.csv", stringsAsFactors = FALSE)

# Extract all reference genotypes
ref_genotypes <- unique(file_info[[1]])

# Get folder list (second column)
folders <- file_info[[2]]

# Function to find the file with 'Result_all' in the folder
find_result_file <- function(folder) {
  files <- list.files(path = folder, pattern = "Result_all.*\\.xls[x]?$", full.names = TRUE)
  if (length(files) == 0) stop(paste("No 'Result_all' file found in:", folder))
  return(files[1])
}

# Get list of file paths from folders
file_paths <- sapply(folders, find_result_file)

# Identify final (last) file (used as-is)
final_file <- tail(file_paths, 1)
df_final <- read_excel(final_file)

# Initialize list to store matching subsets
subset_list <- list()

# Loop through all files except the last one
for (i in 1:(length(file_paths) - 1)) {
  df <- read_excel(file_paths[i])
  
  # Subset rows where genotype column matches any reference genotype
  df_sub <- df[df[[1]] %in% ref_genotypes, ]
  
  # Match common columns with df_final
  common_cols <- intersect(names(df_final), names(df_sub))
  df_sub_matched <- df_sub[, common_cols, drop = FALSE]
  df_final_matched <- df_final[, common_cols, drop = FALSE]
  
  subset_list[[i]] <- df_sub_matched
}

# Combine all subsets and bind to final
combined_subsets <- do.call(rbind, subset_list)
df <- rbind(df_final_matched, combined_subsets)


#********************************************************************
#*P value calculation

# Get all parameter columns (excluding the genotype column)
  param_cols <- names(df)[-c(1:5)]

# Get all genotypes
all_genotypes <- unique(df[[1]])
other_genotypes <- setdiff(all_genotypes, ref_genotypes)

# Initialize an empty list to collect results
pval_list <- list()

# Loop through each genotype (excluding references)
for (gt in other_genotypes) {
  pvals <- c()
  for (param in param_cols) {
    x <- df[df[[1]] %in% ref_genotypes, param, drop = TRUE]
    y <- df[df[[1]] == gt, param, drop = TRUE]
    
    # Make sure we have enough data for the test
    if (length(unique(x)) > 1 && length(unique(y)) > 1) {
      test <- t.test(x, y, var.equal = TRUE)
      pvals <- c(pvals, test$p.value)
    } else {
      pvals <- c(pvals, NA)
    }
  }
  pval_list[[gt]] <- pvals
}


# Create final result data frame
result_df <- as.data.frame(do.call(rbind, pval_list))
colnames(result_df) <- param_cols
result_df <- tibble::rownames_to_column(result_df, var = "Genotype")


#subDir<-readline(prompt = "Enter a folder name to save the results: "); 
#dir.create(file.path(mainDir, subDir))
New_file_location<-file.path(mainDir)


file_neme<-readline(prompt = "Enter a file name to save the results: "); 
file_neme_to_save<-paste0(file_neme,'_PValues.xlsx')

xl_output<-file.path(New_file_location, file_neme_to_save)
# Save the combined dataframe to a new Excel file
write.xlsx(result_df, xl_output, rowNames = FALSE)
