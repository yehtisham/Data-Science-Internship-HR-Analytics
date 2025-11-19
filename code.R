getwd()
# Load necessary library
library(readxl)
library(dplyr)
install.packages("lubridate")
library(lubridate)
library(dplyr)
install.packages("janitor")
library(janitor)

# Define file names
approved_leaves_file <- "Approved Leaves - June 2024.xlsx"
attendance_sheet_file <- "Attendance Sheet - June 2024.xlsx"
consolidated_absent_sheet_file <- "Consolidated Absent Sheet - June 2024.xlsx"
invalid_entries_file <- "Invalid Entries - June 2024.xlsx"
pending_leaves_file <- "Pending Leaves - June 2024.xlsx"

# Read Excel files
approved_leaves <- read_excel(approved_leaves_file)
attendance_sheet <- read_excel(attendance_sheet_file)
consolidated_absent_sheet <- read_excel(consolidated_absent_sheet_file)
invalid_entries <- read_excel(invalid_entries_file)
pending_leaves <- read_excel(pending_leaves_file)

# Display the structure of each dataset
str(approved_leaves)
str(attendance_sheet)
str(consolidated_absent_sheet)
str(invalid_entries)
str(pending_leaves)

# Check the percentage of NA values in each column
na_percentage <- function(df) {
  sapply(df, function(x) sum(is.na(x)) / length(x) * 100)
}

na_percentage(approved_leaves)
na_percentage(attendance_sheet)
na_percentage(consolidated_absent_sheet)
na_percentage(invalid_entries)
na_percentage(pending_leaves)


# Clean column names
approved_leaves <- clean_names(approved_leaves)
attendance_sheet <- clean_names(attendance_sheet)
consolidated_absent_sheet <- clean_names(consolidated_absent_sheet)
invalid_entries <- clean_names(invalid_entries)
pending_leaves <- clean_names(pending_leaves)

# Display cleaned column names
colnames(approved_leaves)
colnames(attendance_sheet)
colnames(consolidated_absent_sheet)
colnames(invalid_entries)
colnames(pending_leaves)



#invalid entries
# Remove rows with any NA values
invalid_entries_clean4 <- na.omit(invalid_entries)
head(invalid_entries_clean4)

# Clean the column names to ensure they are valid
invalid_entries_clean4 <- janitor::clean_names(invalid_entries_clean4)

# Remove the first column
invalid_entries_clean4 <- invalid_entries_clean4 %>%
  select(-1)  # This removes the first column

# Set the first row as column names
colnames(invalid_entries_clean4) <- invalid_entries_clean4[1, ]

# Remove the first row
invalid_entries_clean4 <- invalid_entries_clean4[-1, ]

# Function to convert decimal time to hours and minutes
decimal_to_hm <- function(decimal_time) {
  hours <- floor(decimal_time * 24)
  minutes <- floor((decimal_time * 24 - hours) * 60)
  sprintf("%02d:%02d", hours, minutes)
}

# Convert At Date to Date format and Time In and Time Out from decimal to HH:MM
invalid_entries_clean4 <- invalid_entries_clean4 %>%
  mutate(
    `Time In` = decimal_to_hm(as.numeric(`Time In`)),   # Convert Time In to HH:MM format
    `Time Out` = decimal_to_hm(as.numeric(`Time Out`)), # Convert Time Out to HH:MM format
    `At Date` = as.Date(as.numeric(`At Date`), origin = "1899-12-30")  # Convert At Date to Date format
  )

# Impute missing Time In and Time Out values with "00:00"
invalid_entries_clean4 <- invalid_entries_clean4 %>%
  mutate(
    `Time In` = ifelse(is.na(`Time In`) | `Time In` == "NA:NA", "00:00", `Time In`),
    `Time Out` = ifelse(is.na(`Time Out`) | `Time Out` == "NA:NA", "00:00", `Time Out`)
  )


#approved leaves
# Remove rows with any NA values
approved_leaves_clean1 <- na.omit(approved_leaves)
head(approved_leaves_clean1)

# Remove the first column
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  select(-1)  # This removes the first column

# Set the first row as column names
colnames(approved_leaves_clean1) <- approved_leaves_clean1[1, ]

# Remove the first row
approved_leaves_clean1 <- approved_leaves_clean1[-1, ]


# Convert From Date from decimal to date format
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  mutate(
    `From Date` = decimal_to_date(as.numeric(`From Date`))  # Convert to date format
  )
# Convert From Date column to "MM/D# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert From Date column to "MM/DD/YYYY" format
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  mutate(
    `From Date` = format(`From Date`, format = "%m/%d/%Y")
  )

# Convert To Date from decimal to date format
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  mutate(
    `To Date` = decimal_to_date(as.numeric(`To Date`))  # Convert to date format
  )
# Convert To Date column to "MM/D# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert To Date column to "MM/DD/YYYY" format
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  mutate(
    `To Date` = format(`To Date`, format = "%m/%d/%Y")
  )

# Function to convert decimal dates to date-time format
decimal_to_datetime <- function(decimal_date) {
  as.POSIXct(decimal_date * 86400, origin = "1899-12-30", tz = "UTC")
}

# Convert Entry Date from decimal to date-time format and format to "MM/DD/YYYY HH:MM"
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  mutate(
    `Entry Date` = decimal_to_datetime(as.numeric(`Entry Date`)),
    `Entry Date` = format(`Entry Date`, format = "%m/%d/%Y %H:%M")
  )



#pending leaves
## Clean column names
pending_leaves <- janitor::clean_names(pending_leaves)


# Remove rows with any NA values
pending_leaves_clean10 <- na.omit(pending_leaves)

# Remove the first column
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  select(-1)  # This removes the first column

# Set the first row as column names
colnames(pending_leaves_clean10) <- pending_leaves_clean10[1, ]

# Remove the first row
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  slice(-1)

# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert DOJ from decimal to date format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    DOJ = decimal_to_date(as.numeric(DOJ))  # Convert DOJ to date format
  )
# Convert DOJ column to "MM/D# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert DOJ column to "MM/DD/YYYY" format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(DOJ = format(DOJ, format = "%m/%d/%Y"))


# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert Leave Start From from decimal to date format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    `Leave Start From` = decimal_to_date(as.numeric(`Leave Start From`))
  )

# Convert Leave Start From column to "MM/DD/YYYY" format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    `Leave Start From` = format(`Leave Start From`, format = "%m/%d/%Y")
  )

# Convert Leave Ends On from decimal to date format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    `Leave Ends On` = decimal_to_date(as.numeric(`Leave Ends On`))
  )

# Convert Leave Ends On column to "MM/DD/YYYY" format
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    `Leave Ends On` = format(`Leave Ends On`, format = "%m/%d/%Y")
  )


# Function to convert decimal dates to date-time format
decimal_to_datetime <- function(decimal_date) {
  as.POSIXct(decimal_date * 86400, origin = "1899-12-30", tz = "UTC")
}

# Convert Leave Posting Date from decimal to date-time format and format to "MM/DD/YYYY HH:MM"
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  mutate(
    `Leave Posting Date` = decimal_to_datetime(as.numeric(`Leave Posting Date`)),
    `Leave Posting Date` = format(`Leave Posting Date`, format = "%m/%d/%Y %H:%M")
  )

#
#consolidated absent sheet

# Assuming consolidated_absent_sheet is your dataframe
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet %>%
  select(-starts_with("x"))

# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert Date Of Joining from decimal to date format and then format to "MM/DD/YYYY"
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  mutate(
    `date_of_joining` = decimal_to_date(as.numeric(`date_of_joining`)),
    `date_of_joining` = format(`date_of_joining`, format = "%m/%d/%Y")
  )


# Remove the first column from the dataframe
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  select(-1)

#
#attendance sheet

# Remove rows with NA values using na.omit()
attendance_sheet_clean2 <- na.omit(attendance_sheet)

# Remove the first column from the dataframe
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  select(-1)

# Rename the first column to 'Emp Index'
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
rename(`Emp Index` = colnames(attendance_sheet_clean2)[1])%>%
rename(`Employee Name` = colnames(attendance_sheet_clean2)[2])%>%
rename(`Emp ID` = colnames(attendance_sheet_clean2)[3])%>%
rename(`Position Name` = colnames(attendance_sheet_clean2)[4])%>%
rename(`DOJ` = colnames(attendance_sheet_clean2)[5])


# Define the new column names
new_column_names <- paste0(1:30, " June")

# Rename columns x7 to x36
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  rename_with(~ new_column_names, .cols = x7:x36)

# Define the new column names for x37 to x55
new_column_names <- c(
  "Total Days", "Holidays", "Attended", "Unapproved Leaves", "Approved Leaves", 
  "Absent", "Partial Absent", "Total Gazetted", "Present On WD", "Present On OD", 
  "Present On GD", "Present On Lv", "Overtime WD", "Overtime OD", "Overtime GD", 
  "Irregular", "Leave Deduction", "Late", "Invalid"
)

# Assuming attendance_sheet_clean is your dataframe
# Rename columns x37 to x55
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  rename_with(~ new_column_names, .cols = x37:x55)

# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert DOJ from decimal to date format
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  mutate(
    DOJ = decimal_to_date(as.numeric(DOJ))  # Convert DOJ to date format
  )
# Convert DOJ column to "MM/D# Function to convert decimal dates to date format
decimal_to_date <- function(decimal_date) {
  as.Date(decimal_date, origin = "1899-12-30")
}

# Convert DOJ column to "MM/DD/YYYY" format
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  mutate(DOJ = format(DOJ, format = "%m/%d/%Y"))


# Function to convert decimal time to two distinct times
decimal_to_two_times <- function(decimal_time) {
  # First time calculation
  hours1 <- floor(decimal_time * 24)
  minutes1 <- floor((decimal_time * 24 - hours1) * 60)
  time1 <- sprintf("%02d:%02d", hours1, minutes1)
  
  # Assuming a fixed 8-hour workday for the second time
  hours2 <- (hours1 + 8) %% 24
  time2 <- sprintf("%02d:%02d", hours2, minutes1)
  
  paste(time1, time2, sep = "\n")
}

# Convert decimal times to two times in single cells for columns from "1 June" to "30 June"
attendance_sheet_clean2 <- attendance_sheet_clean2 %>%
  mutate(across(`1 June`:`30 June`, ~ ifelse(!is.na(as.numeric(.)),
                                             decimal_to_two_times(as.numeric(.)),
                                             .)))
#
#Renaming Coloumns

# Rename multiple columns in consolidated_absent_sheet_clean1
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  rename(`Employee Name` = employee_name,
         `Emp ID` = emp_id)

# Rename the column in invalid_entries_clean4
invalid_entries_clean4 <- invalid_entries_clean4 %>%
  rename(`Employee Name` = `Emp Name`)

# Rename the column in approved_leaves_clean1
approved_leaves_clean1 <- approved_leaves_clean1 %>%
  rename(`Employee Name` = `Emp Name`)

# Rename the column in consolidated_absent_sheet_clean1
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  rename(DOJ = date_of_joining)
         
# Rename the column in consolidated_absent_sheet_clean1
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  rename(Location = location)

# Rename the column in consolidated_absent_sheet_clean1
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  rename(Department = department)

# Rename the column in consolidated_absent_sheet_clean1
consolidated_absent_sheet_clean1 <- consolidated_absent_sheet_clean1 %>%
  rename(`Position Name` = position_name)

#
##Merging and Creating Master Sheet

library(tidyr)
library(ggplot2)

# Rename the column in pending_leaves_clean10
pending_leaves_clean10 <- pending_leaves_clean10 %>%
  rename(`Emp ID` = `Emp Id`)


# Pivot the attendance data to long format
attendance_long <- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status")

# Inspect unique values in the Status column
unique_status_values <- unique(attendance_long$Status)

# Display the unique values
print(unique_status_values)


# Refine criteria based on unique status values (Example criteria, to be adjusted based on actual data)
# Aggregate attendance data
attendance_summary2<- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status") %>%
  group_by(`Emp Index`) %>%
  summarise(
    Total_Days_Present = sum(grepl("\\d{2}:\\d{2}\n\\d{2}:\\d{2}", Status), na.rm = TRUE),
    Total_Days_Absent = sum(Status == "Absent", na.rm = TRUE),
    Total_Holidays_OffDays = sum(Status %in% c("Off Day", "Sunday", "Eid ul Azha Holidays", "Off Day Youm E Takbeer", "Off Day 1st Of Jun", "Off Day 2nd Of June", "Off Day 13th June", "Off Day 17th Of Jun", "Off Day Yom-e-Takbeer", "Off Day GH", "Off Day 31st Mar Off", "Off Day 5th Apr Off", "Off Day 19th Apr Off", "Off Day Double Shift 6th May", "Off Day 17th Of May", "Off Day 24th Of May", "Off Day Day Off Of 21st June", "Labor Day", "Labor Day 1st May", "Eid Holidays 4th Day Of Eid Ul Fitr", "Eid ul Azha Holidays 2nd Day Of Eid", "Eid ul Azha Holidays 3rd Day Of Eid", "Eid ul Azha Holidays GH", "Eid Holidays GH"), na.rm = TRUE),
    Total_Leaves = sum(Status %in% c("Casual Leave", "Sick Leave", "Annual Leave", "Compensatory Leave", "Maternity Leave", "Paternity Leave", "Leave Without Pay", "Election", "Pakistan Day Youm-e-Takbeer"), na.rm = TRUE),
    .groups = 'drop'
  )

# Display the summary
head(attendance_summary2)

# Rename the dataset to master_sheet
master_sheet3 <- attendance_summary2

# Display the first few rows of the master sheet to verify
head(master_sheet3)

# Merge master_sheet with attendance_sheet_clean2 to add Emp ID
master_sheet3 <- master_sheet3 %>%
  left_join(attendance_sheet_clean2 %>% select(`Emp Index`, `Emp ID`), 
            by = c("Emp Index"))

# Display the first few rows of the updated master sheet to verify
head(master_sheet3)

# Reorder columns to move Emp ID to the start
master_sheet3 <- master_sheet3 %>%
  select(`Emp ID`, everything())


# Merge master_sheet with attendance_sheet_clean2 to add Late
master_sheet3 <- master_sheet3 %>%
  left_join(attendance_sheet_clean2 %>% select(`Emp Index`, Late), 
            by = c("Emp Index"))

# Display the first few rows of the updated master sheet to verify
head(master_sheet3)

# Step 1: Aggregate the count of invalid entries per employee in invalid_entries_clean4
invalid_entries_count <- invalid_entries_clean4 %>%
  group_by(`Emp Index`) %>%
  summarise(Invalid_Entries = n(), .groups = 'drop')

# Step 2: Merge the invalid entries count with master_sheet
master_sheet3 <- master_sheet3 %>%
  left_join(invalid_entries_count, by = "Emp Index") %>%
  mutate(Invalid_Entries = ifelse(is.na(Invalid_Entries), 0, Invalid_Entries))

# Display the first few rows of the updated master sheet to verify
head(master_sheet3)


# Function to calculate working hours from time string
calculate_working_hours <- function(time_str) {
  if (grepl("\n", time_str)) {
    times <- strsplit(time_str, "\n")[[1]]
    time_in <- hms::as_hms(paste0(times[1], ":00"))
    time_out <- hms::as_hms(paste0(times[2], ":00"))
    duration <- as.numeric(difftime(time_out, time_in, units = "hours"))
    # Adjust for overnight shifts (assuming maximum 24 hours shifts)
    if (duration < 0) {
      duration <- duration + 24
    }
    return(duration)
  } else {
    return(0)
  }
}

# Pivot longer to get daily data
attendance_long2 <- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status")

# Calculate daily working hours
attendance_long2 <- attendance_long2 %>%
  mutate(Working_Hours = sapply(Status, calculate_working_hours))

# Summarise total working hours per employee
working_hours_summary2 <- attendance_long2 %>%
  group_by(`Emp Index`) %>%
  summarise(Total_Working_Hours = sum(Working_Hours, na.rm = TRUE), .groups = 'drop')

# Merge total working hours with master_sheet
master_sheet3 <- master_sheet3 %>%
  left_join(working_hours_summary2, by = c("Emp Index"))

# Display the first few rows of the updated master sheet to verify
head(master_sheet3)


install.packages("purr")
library(purrr)
library(dplyr)
library(tidyr)

# Merge master_sheet2 with attendance_sheet_clean2 to add or update Employee Name
master_sheet3 <- master_sheet3 %>%
  left_join(attendance_sheet_clean2 %>% select(`Emp Index`, `Employee Name`), by = "Emp Index")

# Move the Employee Name column to the start
master_sheet3 <- master_sheet3 %>%
  select(`Emp ID`, `Employee Name`, everything())


##
#correcting attendance summary

# Add Holidays and Attended to attendance_summary2
attendance_summary2 <- attendance_summary2 %>%
  left_join(attendance_sheet_clean2 %>% select(`Emp Index`, Holidays, Attended), by = "Emp Index")

# Display the first few rows to verify
head(attendance_summary2)

# Add Holidays and Attended to master_sheet3
master_sheet3 <- master_sheet2 %>%
  left_join(attendance_summary2 %>% select(`Emp Index`, Holidays, Attended), by = "Emp Index")


install.packages("hms")
library(hms)

library(dplyr)
library(tidyr)
library(hms)

# Create the initial master_sheet4 with Holidays and Attended columns
master_sheet4 <- attendance_sheet_clean2 %>%
  select(`Emp Index`, `Employee Name`, `Emp ID`, Holidays, Attended)

# Function to calculate working hours from time string
calculate_working_hours <- function(time_str) {
  if (grepl("\n", time_str)) {
    times <- strsplit(time_str, "\n")[[1]]
    time_in <- hms::as_hms(paste0(times[1], ":00"))
    time_out <- hms::as_hms(paste0(times[2], ":00"))
    duration <- as.numeric(difftime(time_out, time_in, units = "hours"))
    # Adjust for overnight shifts (assuming maximum 24 hours shifts)
    if (duration < 0) {
      duration <- duration + 24
    }
    return(duration)
  } else {
    return(0)
  }
}

# Pivot longer to get daily data
attendance_long <- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status")

# Calculate daily working hours and invalid entries
attendance_long <- attendance_long %>%
  mutate(
    Working_Hours = sapply(Status, calculate_working_hours),
    Invalid_Entries = ifelse(Status == "Invalid", 1, 0),
    Absent = ifelse(Status == "Absent", 1, 0)
  )

# Summarise total working hours, invalid entries, and absences per employee
attendance_summary <- attendance_long %>%
  group_by(`Emp Index`) %>%
  summarise(
    Total_Working_Hours = sum(Working_Hours, na.rm = TRUE),
    Invalid_Entries = sum(Invalid_Entries, na.rm = TRUE),
    Total_Absent = sum(Absent, na.rm = TRUE),
    .groups = 'drop'
  )

# Merge the summary with master_sheet4
master_sheet4 <- master_sheet4 %>%
  left_join(attendance_summary, by = "Emp Index")

# Display the first few rows of the updated master_sheet4 to verify
head(master_sheet4)

# Add the leaves column from consolidated_absent_sheet_clean1
master_sheet4 <- master_sheet4 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, leaves), by = "Emp Index")

# Rename the leaves column to Total Leaves in master_sheet4
master_sheet4 <- master_sheet4 %>%
  rename(`Total Leaves` = leaves)

# Reorder columns to place Attended before Holidays
master_sheet4 <- master_sheet4 %>%
  select(`Emp ID`, `Emp Index`, `Employee Name`, Attended, Holidays, Total_Absent, Total_Working_Hours, everything())

# Add the off_day column from consolidated_absent_sheet_clean1
master_sheet4 <- master_sheet4 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, off_day), by = "Emp Index")

# Rename the off_day column to Off Days
master_sheet4 <- master_sheet4 %>%
  rename(`Off Days` = off_day)
 
# Remove the Holidays column from master_sheet4
master_sheet4 <- master_sheet4 %>%
  select(-Holidays)

# Add the holiday column from consolidated_absent_sheet_clean1 to master_sheet4
master_sheet4 <- master_sheet4 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, holiday), by = "Emp Index")

# Rename holiday to Total Holidays
master_sheet4 <- master_sheet4 %>%
  rename(`Total Holidays` = holiday)

# Add the 'late' column from consolidated_absent_sheet_clean1 to master_sheet4
master_sheet4 <- master_sheet4 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, late), by = "Emp Index")

# Rename the 'late' column to 'Late' in master_sheet4
master_sheet4 <- master_sheet4 %>%
  rename(Late = late)

# Aggregate Leave Type and Leave Days by Emp Index
leave_aggregated <- approved_leaves_clean1 %>%
  group_by(`Emp Index`, `Leave Type`) %>%
  summarise(Total_Leave_Days = sum(as.numeric(`Leave Days`), na.rm = TRUE), .groups = 'drop') %>%
  pivot_wider(names_from = `Leave Type`, values_from = Total_Leave_Days, values_fill = list(Total_Leave_Days = 0))

# Merge aggregated leave data with master_sheet4
master_sheet4 <- master_sheet4 %>%
  left_join(leave_aggregated, by = "Emp Index")

# Select Emp Index and Position Name from attendance_sheet_clean2
position_data <- attendance_sheet_clean2 %>%
  select(`Emp Index`, `Position Name`)

# Merge Position Name data with master_sheet4
master_sheet4 <- master_sheet4 %>%
  left_join(position_data, by = "Emp Index")

# Reorder columns to move Position Name after Employee Name
master_sheet4 <- master_sheet4 %>%
  select(`Emp ID`, `Emp Index`, `Employee Name`, `Position Name`, everything())

# Rename the specified columns in master_sheet4
master_sheet4 <- master_sheet4 %>%
  rename(
    Absent = Total_Absent,
    `Working Hours` = Total_Working_Hours,
    `Invalid Entries` = Invalid_Entries,
    Holidays = `Total Holidays`
  )

# Check for NA values in the entire dataset
na_summary <- colSums(is.na(master_sheet3))
print(na_summary)

# Alternatively, to view the rows with NA values
na_rows <- master_sheet3[!complete.cases(master_sheet3), ]
print(na_rows)


# Remove specified columns from master_sheet4
master_sheet4 <- master_sheet4 %>%
  select(-`Total Leaves`, -`Off Days`, -`Holidays`, -`Late`)

# Merge desired columns from consolidated_absent_sheet_clean1
master_sheet4 <- master_sheet4 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, leaves, off_day, holiday, late), by = "Emp Index")

# Rename the specified columns in master_sheet4
master_sheet4 <- master_sheet4 %>%
  rename(
    `Total Leaves` = leaves,
    `Off Days` = off_day,
    `Holidays` = holiday,
    `Late` = late
  )


# Step 1: Create a dataset na_rows with rows containing NA values
na_rows2 <- master_sheet4[!complete.cases(master_sheet4), ]

# Step 2: Remove rows containing NA values from master_sheet4
master_sheet4 <- master_sheet4[complete.cases(master_sheet4), ]

# Step 3: Remove all columns from na_summary except specified ones
required_columns <- c("Emp ID", "Emp Index", "Employee Name", "Position Name", "Attended", "Absent")
na_summary <- na_rows[, required_columns]

# Step 4: Create a new column called Without Schedule with values "Yes"
na_summary$`Without Schedule` <- "Yes"

# Step 5: Merge na_summary into master_sheet4 and set "No" for rows not in na_summary
master_sheet4 <- merge(master_sheet4, na_summary, by = required_columns, all.x = TRUE)

# Set "No" for Without Schedule in rows not in na_summary
master_sheet4$`Without Schedule`[is.na(master_sheet4$`Without Schedule`)] <- "No"

# Display the first few rows of the updated master sheet to verify
head(master_sheet4)



##
# creating master sheet again for missing NA values


# Create the initial master_sheet7 with Holidays and Attended columns
master_sheet7 <- attendance_sheet_clean2 %>%
  select(`Emp Index`, `Employee Name`, `Emp ID`, Holidays, Attended)

# Function to calculate working hours from time string
calculate_working_hours <- function(time_str) {
  if (grepl("\n", time_str)) {
    times <- strsplit(time_str, "\n")[[1]]
    time_in <- hms::as_hms(paste0(times[1], ":00"))
    time_out <- hms::as_hms(paste0(times[2], ":00"))
    duration <- as.numeric(difftime(time_out, time_in, units = "hours"))
    # Adjust for overnight shifts (assuming maximum 24 hours shifts)
    if (duration < 0) {
      duration <- duration + 24
    }
    return(duration)
  } else {
    return(0)
  }
}

# Pivot longer to get daily data
attendance_long <- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status")

# Calculate daily working hours and invalid entries
attendance_long <- attendance_long %>%
  mutate(
    Working_Hours = sapply(Status, calculate_working_hours),
    Invalid_Entries = ifelse(Status == "Invalid", 1, 0),
    Absent = ifelse(Status == "Absent", 1, 0)
  )

# Summarise total working hours, invalid entries, and absences per employee
attendance_summary <- attendance_long %>%
  group_by(`Emp Index`) %>%
  summarise(
    Total_Working_Hours = sum(Working_Hours, na.rm = TRUE),
    Invalid_Entries = sum(Invalid_Entries, na.rm = TRUE),
    Total_Absent = sum(Absent, na.rm = TRUE),
    .groups = 'drop'
  )

# Merge the summary with master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(attendance_summary, by = "Emp Index")

# Add the leaves column from consolidated_absent_sheet_clean1
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, leaves), by = "Emp Index")

# Rename the leaves column to Total Leaves in master_sheet7
master_sheet7 <- master_sheet7 %>%
  rename(`Total Leaves` = leaves)

# Reorder columns to place Attended before Holidays
master_sheet7 <- master_sheet7 %>%
  select(`Emp ID`, `Emp Index`, `Employee Name`, Attended, Holidays, Total_Absent, Total_Working_Hours, everything())

# Add the off_day column from consolidated_absent_sheet_clean1
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, off_day), by = "Emp Index")

# Rename the off_day column to Off Days
master_sheet7 <- master_sheet7 %>%
  rename(`Off Days` = off_day)

# Remove the Holidays column from master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(-Holidays)

# Add the holiday column from consolidated_absent_sheet_clean1 to master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, holiday), by = "Emp Index")

# Rename holiday to Total Holidays
master_sheet7 <- master_sheet7 %>%
  rename(`Total Holidays` = holiday)

# Add the 'late' column from consolidated_absent_sheet_clean1 to master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, late), by = "Emp Index")

# Rename the 'late' column to 'Late' in master_sheet7
master_sheet7 <- master_sheet7 %>%
  rename(Late = late)

# Aggregate Leave Type and Leave Days by Emp Index
leave_aggregated <- approved_leaves_clean1 %>%
  group_by(`Emp Index`, `Leave Type`) %>%
  summarise(Total_Leave_Days = sum(as.numeric(`Leave Days`), na.rm = TRUE), .groups = 'drop') %>%
  pivot_wider(names_from = `Leave Type`, values_from = Total_Leave_Days, values_fill = list(Total_Leave_Days = 0))

# Merge aggregated leave data with master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(leave_aggregated, by = "Emp Index")

# Select Emp Index and Position Name from attendance_sheet_clean2
position_data <- attendance_sheet_clean2 %>%
  select(`Emp Index`, `Position Name`)

# Merge Position Name data with master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(position_data, by = "Emp Index")

# Reorder columns to move Position Name after Employee Name
master_sheet7 <- master_sheet7 %>%
  select(`Emp ID`, `Emp Index`, `Employee Name`, `Position Name`, everything())

# Rename the specified columns in master_sheet7
master_sheet7 <- master_sheet7 %>%
  rename(
    Absent = Total_Absent,
    `Working Hours` = Total_Working_Hours,
    `Invalid Entries` = Invalid_Entries,
    Holidays = `Total Holidays`
  )

# Display the first few rows of the updated master_sheet7 to verify
head(master_sheet7)

# Find the position of the 'Late' column
late_position <- which(names(master_sheet7) == "Late")

# Select all columns up to (but not including) the 'Late' column
master_sheet7 <- master_sheet7 %>%
  select(1:(late_position - 1))

# Add the 'late' column from consolidated_absent_sheet_clean1 to master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, late), by = "Emp Index")

# Rename the 'late' column to 'Late'
master_sheet7 <- master_sheet7 %>%
  rename(Late = late)

# Check for NA values in the entire dataset
na_summary <- colSums(is.na(master_sheet7))
print(na_summary)

# Alternatively, to view the rows with NA values
na_rows <- master_sheet7[!complete.cases(master_sheet7), ]
print(na_rows)


# Step 1: Create a dataset na_rows with rows containing NA values
na_rows2 <- master_sheet7[!complete.cases(master_sheet7), ]

# Step 2: Remove all columns from na_rows except specified ones
required_columns <- c("Emp ID", "Emp Index", "Employee Name", "Position Name", "Attended", "Absent")
na_rows2 <- na_rows2[, required_columns]

# Step 3: Create a new column called Without Schedule with values "Yes"
na_rows2$`Without Schedule` <- "Yes"

master_sheet7 <- master_sheet7 %>%
  left_join(na_rows, by = c("Emp ID", "Emp Index", "Employee Name", "Position Name", "Attended", "Absent"))

# Set "No" for Without Schedule in rows not in na_rows
master_sheet7$`Without Schedule`[is.na(master_sheet7$`Without Schedule`)] <- "No"

# Remove the specified columns from master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(-c(`Working Hours.x`, `Working Hours.y`, `Invalid Entries.x`, `Invalid Entries.y`, 
            `Total Leaves.x`, `Total Leaves.y`, `Off Days.x`, `Off Days.y`, 
            `Holidays.x`, `Holidays.y`, `Late.x`, `Late.y`))

# Add the specified columns from consolidated_absent_sheet_clean1 to master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(consolidated_absent_sheet_clean1 %>% select(`Emp Index`, leaves, off_day, holiday, late), by = "Emp Index")

# Rename the columns in master_sheet7
master_sheet7 <- master_sheet7 %>%
  rename(
    `Total Leaves` = leaves,
    `Off Days` = off_day,
    Holidays = holiday,
    Late = late
  )

# Step 1: Aggregate the count of invalid entries per employee in invalid_entries_clean4
invalid_entries_count <- invalid_entries_clean4 %>%
  group_by(`Emp Index`) %>%
  summarise(`Invalid Entries` = n(), .groups = 'drop')

# Step 2: Merge the invalid entries count with master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(invalid_entries_count, by = "Emp Index") %>%
  mutate(`Invalid Entries` = ifelse(is.na(`Invalid Entries`), 0, `Invalid Entries`))



# Function to calculate working hours from time string
calculate_working_hours <- function(time_str) {
  if (grepl("\n", time_str)) {
    times <- strsplit(time_str, "\n")[[1]]
    time_in <- hms::as_hms(paste0(times[1], ":00"))
    time_out <- hms::as_hms(paste0(times[2], ":00"))
    duration <- as.numeric(difftime(time_out, time_in, units = "hours"))
    # Adjust for overnight shifts (assuming maximum 24 hours shifts)
    if (duration < 0) {
      duration <- duration + 24
    }
    return(duration)
  } else {
    return(0)
  }
}

# Pivot longer to get daily data
attendance_long5 <- attendance_sheet_clean2 %>%
  pivot_longer(cols = `1 June`:`30 June`, names_to = "Date", values_to = "Status")

# Calculate daily working hours
attendance_long5 <- attendance_long5 %>%
  mutate(Working_Hours = sapply(Status, calculate_working_hours))

# Summarise total working hours per employee
working_hours_summary8 <- attendance_long5 %>%
  group_by(`Emp Index`) %>%
  summarise(`Total Working Hours` = sum(Working_Hours, na.rm = TRUE), .groups = 'drop')

# Merge total working hours with master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(working_hours_summary8, by = "Emp Index")

# Display the first few rows of the updated master_sheet7 to verify
head(master_sheet7)

# Check for NA values in the entire dataset
na_summary <- colSums(is.na(master_sheet7))
print(na_summary)

# Alternatively, to view the rows with NA values
na_rows <- master_sheet7[!complete.cases(master_sheet7), ]
print(na_rows)


# Select the relevant columns from attendance_sheet_clean2
leave_columns <- attendance_sheet_clean2 %>%
  select(`Emp Index`, `Approved Leaves`, `Unapproved Leaves`)

# Merge the leave columns into master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(leave_columns, by = "Emp Index")

# Group by Emp Index and concatenate the leave types
leave_types <- approved_leaves_clean1 %>%
  group_by(`Emp Index`) %>%
  summarise(Leave_Types = paste(unique(`Leave Type`), collapse = ", "), .groups = 'drop')

# Merge the concatenated leave types into master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(leave_types, by = "Emp Index")

# Display the first few rows of the updated master sheet to verify
head(master_sheet7)

# Step 3: Replace NA values in the Leave_Types column with "None"
master_sheet7 <- master_sheet7 %>%
  mutate(Leave_Types = ifelse(is.na(Leave_Types), "None", Leave_Types))

# Rename the Leave_Types column to Type Of Approved Leaves
master_sheet7 <- master_sheet7 %>%
  rename(`Type Of Approved Leaves` = Leave_Types)

#Organizing Columns
# Organize columns in master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(
    `Emp ID`, `Emp Index`, `Employee Name`, `Position Name`, 
    Attended, Absent, `Total Working Hours`, `Invalid Entries`, 
    `Total Leaves`, `Off Days`, Holidays, Late,`Approved Leaves`, `Unapproved Leaves`,
    `Type Of Approved Leaves` , `Without Schedule`
  )

# Convert Attended and Absent columns to numeric
master_sheet7 <- master_sheet7 %>%
  mutate(
    Attended = as.numeric(Attended),
    Absent = as.numeric(Absent)
  )

# Add a new column named Total Assigned Duties by summing Attended and Absent
master_sheet7 <- master_sheet7 %>%
  mutate(`Total Assigned Duties` = Attended + Absent)

# Step 2: Reorder columns to place Total Assigned Duties after Absent
master_sheet7 <- master_sheet7 %>%
  select(
    `Emp ID`, `Emp Index`, `Employee Name`, `Position Name`,
    Attended, Absent, `Total Assigned Duties`, 
    everything()
  )

# Remove the column Type Of Approved Leaves from master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(-`Type Of Approved Leaves`)

# Rename the column Without Schedule to Schedule
master_sheet7 <- master_sheet7 %>%
  rename(`Schedule` = `Without Schedule`)

# Change all "Yes" to "No" and all "No" to "Yes" in the Schedule column
master_sheet7 <- master_sheet7 %>%
  mutate(Schedule = ifelse(Schedule == "Yes", "No", ifelse(Schedule == "No", "Yes", Schedule)))


master_sheet7<- master_sheet7 %>% 
  left_join(leave_types, by = "Emp Index")

# Remove the Leave_Types column from master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(-Leave_Types) 


# Aggregate the leave data by Emp Index and Leave Type
leave_aggregated2 <- approved_leaves_clean1 %>%
  group_by(`Emp Index`, `Leave Type`) %>%
  summarise(Total_Leave_Days = sum(as.numeric(`Leave Days`), na.rm = TRUE), .groups = 'drop')

# Pivot the data to create separate columns for each leave type
leave_pivot1 <- leave_aggregated2 %>%
  pivot_wider(names_from = `Leave Type`, values_from = Total_Leave_Days, values_fill = list(Total_Leave_Days = 0))


# Ensure the Half Day-Casual column is numeric
leave_pivot1 <- leave_pivot1 %>%
  mutate(`Half-Day Casual` = as.numeric(`Half-Day Casual`))

# Sum the values in the column Half-Day Casual
half_day_casual_sum <- sum(leave_pivot1$`Half-Day Casual`, na.rm = TRUE)

# Display the sum
print(half_day_casual_sum)

# Remove the Half-Day Casual column
leave_pivot1 <- leave_pivot1 %>%
  select(-`Half-Day Casual`)

# Ensure the column names match for the join
leave_pivot1 <- leave_pivot1 %>%
  rename(`Emp Index` = `Emp Index`)

# Merge the columns from leave_pivot1 into master_sheet7
master_sheet7 <- master_sheet7 %>%
  left_join(leave_pivot1, by = "Emp Index")

# Count the number of NA values in each column of master_sheet7
na_summary <- sapply(master_sheet7, function(x) sum(is.na(x)))

# Display the NA summary
na_summary


# Replace NA values with 0 in the specified columns
columns_to_replace <- c("Annual Leave", "Sick Leave", "Casual Leave", "Compensatory Leave", 
                        "Paternity Leave", "Maternity Leave", "Short Leave Annual", 
                        "Leave Without Pay")

# Replace NA with 0 in the specified columns
master_sheet7[columns_to_replace] <- lapply(master_sheet7[columns_to_replace], function(x) {
  x[is.na(x)] <- 0
  return(x)
})


##
#Visualization
library(ggplot2)
install.packages("reshape2")
library(reshape2)

# If you don't have a date column and want to add a placeholder month column manually
# For example, if the entire dataset is for June 2024
master_sheet7 <- master_sheet7 %>%
  mutate(Month = "June 2024")

# Remove the Month column from master_sheet7
master_sheet7 <- master_sheet7 %>%
  select(-Month)


# Export master_sheet7 to a CSV file
write.csv(master_sheet7, file = "master_sheet7.csv", row.names = FALSE)


