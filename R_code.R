#install.packages("dplyr", dependencies = TRUE)
if (!requireNamespace("dplyr", quietly = TRUE)) install.packages("dplyr")
if (!requireNamespace("openxlsx", quietly = TRUE)) install.packages("openxlsx")
library(dplyr)
library(openxlsx)

# Read UTF-16 encoded CSV files
csv_path_1 <- "path/to/first_file.csv"
csv_path_2 <- "path/to/second_file.csv"

if (!file.exists(csv_path_1) || !file.exists(csv_path_2)) {
    stop("One or both CSV files do not exist. Please check the file paths.")
}

v1 <- read.csv(csv_path_1, fileEncoding = "UTF-16", stringsAsFactors = FALSE, check.names = FALSE, header = TRUE)
v2 <- read.csv(csv_path_2, fileEncoding = "UTF-16", stringsAsFactors = FALSE, check.names = FALSE, header = TRUE)

# Check the structure of the data
cat("Structure of v1:\n")
str(v1)
cat("Structure of v2:\n")
str(v2)

diff_v1 <- anti_join(v1, v2, by = intersect(names(v1), names(v2)))
diff_v1 <- anti_join(v1, v2)
cat("Rows in version 1 but missing in version 2:\n")
print(diff_v1)
diff_v2 <- anti_join(v2, v1, by = intersect(names(v1), names(v2)))
diff_v2 <- anti_join(v2, v1)
cat("Rows in version 2 but missing in version 1:\n")
print(diff_v2)

diff_v1_case_insensitive <- anti_join(
    v1 %>% mutate(across(everything(), ~tolower(as.character(.)))),
    v2 %>% mutate(across(everything(), ~tolower(as.character(.))))
)
cat("Rows in version 1 but missing in version 2 (case-insensitive):\n")
print(diff_v1_case_insensitive)

diff_v2_case_insensitive <- anti_join(
    v2 %>% mutate(across(everything(), tolower)),
    v1 %>% mutate(across(everything(), tolower))
)
cat("Rows in version 2 but missing in version 1 (case-insensitive):\n")
print(diff_v2_case_insensitive)

# Create a new Excel workbook
wb <- createWorkbook()

# Function to write and format sheets
write_diff_sheet <- function(sheet_name, data) {
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet_name, data, headerStyle = createStyle(textDecoration = "bold"))
    if (nrow(data) > 0) {
        addStyle(wb, sheet_name, createStyle(fgFill = "#FFEB9C"), rows = 2:(nrow(data) + 1), cols = 1:ncol(data), gridExpand = TRUE)
    }
}

# Add summary sheet
addWorksheet(wb, "Summary")
writeData(wb, "Summary", data.frame(
    Comparison = c("Rows only in version 1", "Rows only in version 2"),
    Count = c(nrow(diff_v1), nrow(diff_v2))
))
addStyle(wb, "Summary", createStyle(textDecoration = "bold"), rows = 1, cols = 1:2, gridExpand = TRUE)

cat("Differences exported to 'Differences_Report.xlsx'\n")
write_diff_sheet("Only_in_v1", diff_v1)
write_diff_sheet("Only_in_v2", diff_v2)

# Save the workbook
saveWorkbook(wb, "Differences_Report.xlsx", overwrite = TRUE)
cat("Differences exported to 'Differences_Report.xlsx'\n")
