#install below R packages if not installed on your computer

#install.packages(tidyverse)
#install.packages(readxl)

#loads required packages
library(tidyverse)
library(readxl)

files_list1 <- list("EngineeringHardware02.xlsx", "EngineeringHardware03.xlsx", "EngineeringHardware05.xlsx",
                    "EngineeringHardware06.xlsx", "EngineeringHardware07.xlsx", "EngineeringHardware08.xlsx",
                    "EngineeringHardware09.xlsx","EngineeringHardware10.xlsx", "Sanitary02.xlsx")
files_list2 <- list("EngineeringHardware04.xlsx", "EngineeringHardware11.xlsx") 
df_files1 <- lapply(files_list1, read_xlsx)  # reads xlsx files
df_files2 <- lapply(files_list2, read_xlsx)

#bind data frames in df_files together by rows
company_1 <- bind_rows(df_files1) %>% data.frame()
company_2 <- bind_rows(df_files2) %>% data.frame()

#converts all columns to character
company_1 <- lapply(company_1, as.character) %>% data.frame()
company_2 <- lapply(company_2, as.character) %>% data.frame() 

#removes unwanted columns
company_2$X__1 <- NULL
company_2$X__2 <- NULL
company_2$X__3 <- NULL
company_2$X__4 <- NULL
company_2$X__5 <- NULL
company_2$X__6 <- NULL
company_2$X__7 <- NULL
company_2$X__8 <- NULL



# renames column names
col_names <- c("Id", "Company_Name", "Address_1", "Address_2", "Address_3", "City", "Pin", "State", 
               "Country", "Phone1", "Phone2", "Mobile1", "Mobile2", "Fax", "Email", "Website", "Title", 
               "First_Name", "Last_Name", "Business_Details" )
names(company_1) <- col_names
names(company_2) <- col_names

#binds two dataframes together by rows
company_final <- bind_rows(company_1, company_2)

#unites columns together
unite_1 <- c("Address_1", "Address_2", "Address_3")
unite_2 <- c("Title", "First_Name")
unite_3 <- c('Title_firstname', "Last_Name")
unite_4 <- c("Phone1", "Phone2")
unite_5 <- c("Mobile1", "Mobile2")
company_final <- unite(company_final, Address, unite_1, sep = " ") %>% 
  unite(Title_firstname, unite_2, sep = "") %>%
  unite(Contact_Name, unite_3, sep = " ") %>% 
  unite(Phone, unite_4, sep = "/") %>% 
  unite(Mobile, unite_5, sep = "/")

# remove "NA" from column "phone"
company_final$Phone <- gsub("\\/NA", "", company_final$Phone)
company_final$Mobile <- gsub("\\/NA", "", company_final$Mobile)

#Adds aditional two columns "Designation", "Business_Type"
company_final$Designation <- NA
company_final$Business_Type <- NA

#Rearrange columns
rearrange_col0 <- c(1:13, 15, 14, 16)
company_final <- company_final[, rearrange_col0]

# converts all columns to char
company_final <- lapply(company_final, as.character) %>% data.frame(stringsAsFactors=FALSE)

# for file "EngineeringHardware01.xlsx"
file_1 <- read_xlsx("EngineeringHardware01.xlsx")

new_cols1 <- c("Id", "Country", "Fax", "Contact_Name")
file_1[new_cols1] <- NA

rearrange_columns1 <- c(15, 1:5, 16, 6:9, 17, 10:12, 18, 14, 13)

file_1 <- file_1[, rearrange_columns1]

col1_names <- c("Id", "Company_Name", "Address", "City", "Pin", "State",
                "Country", "Phone1", "Phone2", "Mobile1", "Mobile2", "Fax", "Email1", "Email2", 
                "Website", "Contact_Name", "Business_Details", "Business_Type")

names(file_1) <- col1_names       # renaming column names

# uniting mobile1 and mobile2 etc
unite_6 <- c("Phone1", "Phone2")
unite_7 <- c("Mobile1", "Mobile2")
unite_8 <- c("Email1", "Email2")

file_1<- unite(file_1, Phone, unite_6, sep = "/") %>%
  unite(Mobile, unite_7, sep = "/") %>% 
  unite(Email, unite_8, sep = "/")


# remove "NA" from column "phone"
file_1$Phone <- gsub("\\/NA", "", file_1$Phone)
file_1$Mobile <- gsub("\\/NA", "", file_1$Mobile)
file_1$Email <- gsub("\\/NA", "", file_1$Email)

#Adds aditional two columns "Designation",
file_1$Designation <- NA

#Rearrange columns
rearrange_col1 <- c(1:13, 16, 14, 15)
file_1 <- file_1[, rearrange_col1]

# converts all columns to char
file_1 <- lapply(file_1, as.character) %>% data.frame(stringsAsFactors=FALSE)



# for file "Delivery Lot 1_Sports Equipments.xlsx"
file_2 <- read_xlsx("Delivery Lot 1_Sports Equipments.xlsx")
file_2$X__3 <- NULL
new_cols2 <- c("Country", "Email", "Business_Details", "Business_Type")
file_2[new_cols2] <- NA
col2_names <- c("Id", "Company_Name", "Address", "City", "State", "Pin", "Phone", "Fax", 
                "Website", "Contact_Name", "Designation", "Mobile", "Country", "Email", 
                "Business_Details", "Business_Type")
names(file_2) <- col2_names

rearrange_columns2 <- c(1:4, 6, 5, 13, 7, 12, 8, 14, 9, 10, 11,  15, 16)

file_2 <- file_2[, rearrange_columns2]
file_2 <- lapply(file_2, as.character) %>% data.frame(stringsAsFactors=FALSE)



# for file "Email Id_Insertion.xlsx"
file_3 <- read_xlsx("Email Id_Insertion.xlsx")
file_3$`Assigned to` <- NULL # removes unnessary last columns

new_cols3<- c("Country", "Business_Details", "Business_Type")  #Adds three new empty columns
file_3[new_cols3] <- NA

col3_names <- c("Id", "Company_Name", "Address", "City", "State", "Pin", "Phone", "Fax", "Website", 
                "Contact_Name", "Designation", "Email", "Mobile", "Country", "Business_Details",
                "Business_Type")
names(file_3) <- col3_names

rearrange_columns3 <- c(1:4, 6, 5, 14, 7, 13, 8, 12, 9, 10, 11, 15, 16)

file_3 <- file_3[, rearrange_columns3]
file_3 <- lapply(file_3, as.character) %>% data.frame(stringsAsFactors=FALSE)


# for file "File left to be added.xlsx"
file_4 <- read_xlsx("File left to be added.xlsx", col_names =FALSE)
file_4$X__7 <- NULL # deletes column "X_7"
file_4$X__11 <- NULL

new_cols4 <- c("Country", "Phone", "Email", "Website", "Business_Details", "Business_Type")
file_4[new_cols4] <- NA

col4_names <- c("Id", "Company_Name", "Address", "City", "State", "Pin", "Contact_Name", "Fax", "Designation", "Mobile", "Country", "Phone", "Email", "Website", "Business_Details", "Business_Type")

names(file_4) <- col4_names
rearrange_columns4 <- c(1:4, 6, 5, 11, 12, 10, 8, 13, 14, 7, 9, 15, 16)

file_4 <- file_4[, rearrange_columns4]

file_4 <- lapply(file_4, as.character) %>% data.frame(stringsAsFactors=FALSE)

#for file "hims1.xlsx"
file_5 <- read_xlsx("hims1.xlsx")
new_cols5 <- c("Country", "Business_Details", "Business_Type")

file_5[new_cols5] <- NA
col5_names <- c("Id", "Company_Name", "Address", "City", "State", "Pin", "Phone", "Fax", "Website", "Contact_Name", "Designation", "Email", "Mobile", "Country", "Business_Details", "Business_Type")
names(file_5) <- col5_names

rearrange_columns5 <- c(1:4, 6, 5, 14, 7, 13, 8, 12, 9:11, 15, 16)

file_5 <- file_5[, rearrange_columns5]
file_5 <- lapply(file_5, as.character) %>% data.frame(stringsAsFactors=FALSE)

# for file "Lot 3 _Gym Equipment Manufacturers.xlsx"
file_6 <- read_xlsx("Lot 3 _Gym Equipment Manufacturers.xlsx")
new_cols6 <- c("Country", "Business_Details", "Business_Type")

file_6[new_cols6] <- NA
col6_names <- c("Id", "Company_Name", "Address", "City", "State", 
                "Pin", "Phone", "Fax", "Website", "Contact_Name",
                "Designation", "Email", "Mobile", "Country", "Business_Details", "Business_Type")
names(file_6) <- col6_names

rearrange_columns6 <- c(1:4, 6, 5, 14, 7, 13, 8, 12, 9:11, 15, 16 )

file_6 <- file_6[, rearrange_columns6] # rearrange columns
# converts all columns to char
file_6 <- lapply(file_6, as.character) %>% data.frame(stringsAsFactors=FALSE) 

# for two files "Sports Equipments_Lot 2.xlsx" and "Sports Equipments_Lot 3 (1).xlsx"

files_list2 <- list("Sports Equipments_Lot 2.xlsx", "Sports Equipments_Lot 3 (1).xlsx")
Sports_files <- lapply(files_list2, read_xlsx)
file_7 <- bind_rows(Sports_files) %>% data.frame()
new_cols7 <- c("Country", "Business_Details", "Business_Type")
file_7[new_cols7] <- NA
col7_names <- c("Id", "Company_Name", "Address", "City", "State", "Pin", "Phone", "Fax", 
                "Website", "Contact_Name", "Designation", "Mobile", "Email", "Country", 
                "Business_Details", "Business_Type")
names(file_7) <- col7_names
rearrange_columns7<- c(1:4, 6, 5, 14, 7, 12, 8, 13, 9, 10, 11, 15, 16)

file_7 <- file_7[, rearrange_columns7]
# converts all columns to char
file_7 <- lapply(file_7, as.character) %>% data.frame(stringsAsFactors=FALSE) 


# Binds all dataframes together
data_final <- bind_rows(company_final, file_1, file_2, file_3, file_3, file_4, file_5, file_6, file_7)

#Removes "Id" column
data_final$Id <- NULL

#Removes duplicated data
data_final <- data_final[!duplicated(data_final$Company_Name), ]
data_final <- data_final[!duplicated(data_final$Email), ]

# writes this dataframe as csv file
write.csv(data_final, file = "Jignesh_final.csv")

