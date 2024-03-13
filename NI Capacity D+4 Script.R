settype <- "D+4"
settype2 <- "D+4 Initial" 
settype3 <- "INIT"
settype4 <- "D4"
mm <- 20

startdate <- documentdata %>%  filter(market2 == "CRM" & run_type == "INIT")
startdate <- startdate$bp_start_date %>% unique() %>% as.Date(format = "%Y-%m-%d")
enddate <- documentdata %>% filter(market2 == "CRM" & run_type == "INIT")
enddate <- enddate$bp_end_date %>% unique() %>% as.Date(format = "%Y-%m-%d")

month <- format(startdate, "%b")
year <- format(startdate, "%Y") %>% substr(start = 3, stop = 4)

capacityfilepath <- paste0(accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/")
capacityfiles <- list.files(capacityfilepath, full.names = TRUE)
capacityfiles <- capacityfiles[grepl(pattern = month, x = capacityfiles)]
capacityfiles <- capacityfiles[grepl(pattern = year, x = capacityfiles)]

### STATEMENT

# read in xmlfiles
for (f in 1:days_in_month(startdate)) {
  
  xmlfile <- paste0(accdrive, ":/GeneralAccounts/Settlement/SEMODownloads/", receiveddate,"/", list.files(paste0(accdrive, ":/GeneralAccounts/Settlement/SEMODownloads/", receiveddate, "/"), pattern = paste0("^SS_", PTunit, "_\\d{8}_\\d{8}_CRM_", settype3, "_\\d{8}T\\d{6}\\.XML$")))
  
  if (length(xmlfile) != days_in_month(startdate)) {
    stop("Ensure the correct number of statements exist in SEMODownloads folder before rerunning. Stopping...")
  }
  
  xmlfile <- xmlfile[[f]]
  doc <- read_xml(xmlfile)
  
  # Get statement document number
  v5 <- sapply(getNodeSet(xmlParse(xmlfile), "/REPORT/*"), xmlAttrs)
  statementnumber <- sapply(v5, function(x) ifelse("statement_id" %in% names(x), x["statement_id"], NA))[1]
  rm(v5)
  
  nodes1 <- xml_find_all(doc, xpath = "//REPORT_HEADER")
  df1 <- bind_rows(lapply(nodes1, xml_attrs)) # this will need repeats the lenght of the statement
  df1$settlement_date <- as.Date(df1$settlement_date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  df1$publication_date <- as.Date(df1$publication_date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  df1 <- df1[rep(1:nrow(df1), each = 98), ] # each report has report_header repeated 98 times
  
  
  
  # ================================================================= report summary section ================================================================= #
  # Extract attributes from REPORT_SUMMARY nodes
  nodes_report_detail <- xml_find_all(doc, xpath = "//REPORT_SUMMARY/*")
  df_report_detail <- bind_rows(lapply(nodes_report_detail, xml_attrs))
  
  
  # Access the Sales and Purchases sections in REPORT_SUMMARY
  details <- xml_find_all(doc, "//REPORT_SUMMARY/*")
  
  # Create an empty data frame to store the results
  result_df <- data.frame(
    name = character(),
    date = character(),
    amount = character(),
    stringsAsFactors = FALSE
  )
  
  # Loop through each DETAIL section
  for (detail in details) {
    # Access the market attribute of DETAIL
    name <- xml_attr(detail, "name", default = NA)
    date <- xml_attr(detail, "date", default = NA)
    amount <- xml_attr(detail, "amount", default = NA)
    
    # Append the results to the data frame
    result_df <- rbind.na(result_df, c(name, date, amount))
  }
  
  # Set column names
  colnames(result_df) <- c("name", "date", "amount")
  
  result_df$date <- as.Date(result_df$date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  # =============================================================== report summary section end =============================================================== #
  
  
  # ================================================================= report detail section ================================================================= #
  # Extract attributes from REPORT_DETAIL nodes
  nodes_report_detail <- xml_find_all(doc, xpath = "//REPORT_DETAIL/*")
  df_report_detail <- bind_rows(lapply(nodes_report_detail, xml_attrs))
  
  # Access the RESOURCE sections in REPORT_DETAIL
  resources <- xml_find_all(doc, "//REPORT_DETAIL/RESOURCE")
  
  # Create an empty data frame to store the results
  result_df2 <- data.frame(
    Resource = character(),
    Charge = character(),
    datetime = character(),
    amount = character(),
    stringsAsFactors = FALSE
  )
  
  # Loop through each RESOURCE section
  for (i in seq_along(resources)) {
    # Get the 'name' attribute of the RESOURCE
    resource_name <- xml_attr(resources[[i]], "name", default = NA)
    
    # Access the CHARGE sections within RESOURCE
    charges <- xml_find_all(resources[[i]], ".//CHARGE")
    
    # Loop through each CHARGE section
    for (j in seq_along(charges)) {
      # Get the 'name' attribute of the CHARGE
      charge_name <- xml_attr(charges[[j]], "name", default = NA)
      
      # Access the VALUE elements within CHARGE
      values <- xml_find_all(charges[[j]], ".//VALUE")
      
      # Loop through each VALUE element
      for (k in seq_along(values)) {
        # Access the attributes of VALUE
        datetime <- xml_attr(values[[k]], "datetime", default = NA)
        amount <- xml_attr(values[[k]], "amount", default = NA)
        
        # Append the results to the data frame
        result_df2 <- rbind.na(result_df2, c(resource_name, charge_name, datetime, amount))
      }
    }
  }
  
  # Set column names
  colnames(result_df2) <- c("name2", "name3", "datetime", "amount4")
  # =============================================================== report detail section end =============================================================== #
  
  # Offset start of result_df2 by nrows in result_df
  result_df2 <- insertRows(df = result_df2, r = 1:nrow(result_df))
  
  df <- cbind.na(df1, result_df, result_df2)
  
  writexl::write_xlsx(df, paste0("C:/Users/", Sys.getenv("USERNAME"),"/Downloads//", f, " STATEMENT ", statementnumber, " for ", PTunit, ".xlsx"))
  
}

# combine statements
inputfiles <- list.files(paste0("C:/Users/", Sys.getenv("USERNAME"), "/Downloads"), full.names = TRUE)
inputfiles <- inputfiles[grepl(x = inputfiles, pattern = "*STATEMENT*")]
inputfiles <- inputfiles[grepl(x = inputfiles, pattern = PTunit)]
outputfile <- paste0("C:/Users/", Sys.getenv("USERNAME"),"/Downloads/All ", month, " ", year, " ", settype, " Capacity Statements for ", PTunit, ".xlsx")

dfstatements <- data.frame()

for (file in inputfiles) {
  df <- read_excel(file)
  dfstatements <- rbind(dfstatements, df)
  file.remove(file)
}

writexl::write_xlsx(dfstatements, outputfile)



### REPORT

for (f in 1:days_in_month(startdate)) {
  
  # read in XML file
  xmlfile <- paste0(accdrive, ":/GeneralAccounts/Settlement/SEMODownloads/", receiveddate,"/", list.files(paste0(accdrive, ":/GeneralAccounts/Settlement/SEMODownloads/", receiveddate, "/"), pattern = paste0("^SR_", PTunit, "_\\d{8}_\\d{8}_CRM_", settype3, "_\\d{8}T\\d{6}\\.XML$")))
  
  if (length(xmlfile) != days_in_month(startdate)) {
    stop("Ensure the correct number of reports exist in SEMODownloads folder before rerunning. Stopping...")
  }
  
  xmlfile <- xmlfile[[f]]
  doc <- read_xml(xmlfile)
  
  # Get report id number
  v5 <- sapply(getNodeSet(xmlParse(xmlfile), "/REPORT/*"), xmlAttrs)
  reportnumber <- sapply(v5, function(x) ifelse("statement_id" %in% names(x), x["statement_id"], NA))[1]
  rm(v5)
  
  # Read in report header node
  nodes1 <- xml_find_all(doc, xpath = "//REPORT_HEADER")
  df1 <- bind_rows(lapply(nodes1, xml_attrs))
  df1$settlement_date <- as.Date(df1$settlement_date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  df1$publication_date <- as.Date(df1$publication_date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  df1 <- df1[rep(1:nrow(df1), each = 50), ] # each report has report_header repeated 50 times
  
  # ========================================================== report detail - determinant section ========================================================== #
  # Extract attributes from REPORT_DETAIL nodes
  nodes_detail <- xml_find_all(doc, xpath = "//REPORT_DETAIL/*")
  df_report_detail <- bind_rows(lapply(nodes_detail, xml_attrs))
  
  # Access the DETERMINANT sections in REPORT_DETAIL
  determinants <- xml_find_all(doc, "//REPORT_DETAIL/DETERMINANT")
  
  # Create an empty data frame to store the results
  result_df <- data.frame(
    name = character(),
    unit = character(),
    datetime = character(),
    amount = character(),
    date = character(),
    stringsAsFactors = FALSE
  )
  
  # Loop through each DETERMINANT section
  for (i in seq_along(determinants)) {
    # Get the 'name' attribute of the RESOURCE
    name <- xml_attr(determinants[[i]], "name")
    unit <- xml_attr(determinants[[i]], "unit")
    
    # Access the VALUE sections within DETERMINANT
    values <- xml_find_all(determinants[[i]], ".//VALUE")
    
    # Loop through each VALUE section
    for (j in seq_along(values)) {
      # Get the attributes of the VALUE
      datetime <- xml_attr(values[[j]], "datetime")
      date <- xml_attr(values[[j]], 'date', default = NA)
      amount <- xml_attr(values[[j]], "amount")
      
      # Append the results to the data frame
      result_df <- rbind.na(result_df, c(name, unit, datetime, amount, date))
    }
  }
  
  result_df$date <- as.Date(result_df$date, format = "%Y-%m-%d") %>% format("%d/%m/%Y") %>% as.Date(format = "%d/%m/%Y")
  
  # ======================================================== report detail - determinant section end ======================================================== #
  
  
  
  
  # =========================================================== report detail - resource section =========================================================== #
  # Extract attributes from REPORT_DETAIL nodes
  nodes_detail <- xml_find_all(doc, xpath = "//REPORT_DETAIL/*")
  df_report_detail <- bind_rows(lapply(nodes_detail, xml_attrs))
  
  # Access the RESOURCE sections in REPORT_DETAIL
  resources <- xml_find_all(doc, "//REPORT_DETAIL/RESOURCE")
  
  # Create an empty data frame to store the results
  result_df2 <- data.frame(
    name2 = character(),
    stringsAsFactors = FALSE
  )
  
  # Loop through each RESOURCE section
  for (i in seq_along(resources)) {
    # Get the 'name' attribute of the RESOURCE
    name2 <- xml_attr(resources[[i]], "name")
  }
  
  name2 <- rep.int(name2, 50) %>% as.data.frame() # each report has name2 repeated 50 times
  colnames(name2) <- "name2"
  
  
  
  # ========================================================= report detail - resource section end ========================================================= #
  
  
  # Create one df
  df <- cbind.na(df1, result_df, name2)
  
  # Create excel file
  writexl::write_xlsx(df, paste0("C:/Users/", Sys.getenv("USERNAME"),"/Downloads//", f, " REPORT ", reportnumber, " for ", PTunit, ".xlsx"))
}

# combine reports
inputfiles <- list.files(paste0("C:/Users/", Sys.getenv("USERNAME"), "/Downloads"), full.names = TRUE)
inputfiles <- inputfiles[grepl(x = inputfiles, pattern = "*REPORT*")]
inputfiles <- inputfiles[grepl(x = inputfiles, pattern = PTunit)]
outputfile <- paste0("C:/Users/", Sys.getenv("USERNAME"),"/Downloads/All ", month, " ", year, " ", settype," Capacity Reports for ", PTunit, ".xlsx")

dfreports <- data.frame()

for (file in inputfiles) {
  df <- read_excel(file)
  dfreports <- rbind(dfreports, df)
  file.remove(file)
}

writexl::write_xlsx(dfreports, outputfile)



# create excel file - setup

inputfiles <- list.files(path = paste0("C:/Users/", Sys.getenv("USERNAME"),"/Downloads/"), full.names = TRUE)
inputfiles <- inputfiles[grepl(pattern = "\\.xlsx$", x = inputfiles)]
inputfiles <- inputfiles[grepl(pattern = PTunit, x = inputfiles)]
inputfiles <- inputfiles[grepl(pattern = "Document|Capacity", x = inputfiles)]

month.as.number <- lubridate::month(startdate)

if (nchar(month.as.number) == 1) {
  month.as.number <- paste0("0", month.as.number)
}

# Get file path for output file
outputfile <- paste0(accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/20", year, " ", month.as.number, ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx")
if (file.exists(outputfile)) {
  file.remove(outputfile)
}

options(openxlsx.dateFormat = "dd/mm/yyyy")

# create excel file
mywb <- createWorkbook()

# add sheets
addWorksheet(wb = mywb, sheetName = "CRM Summary", tabColour = '#F4B084')

temp <- read_excel(path = inputfiles[2], sheet = 1)
addWorksheet(mywb, sheet = "CRM STATEMENT", tabColour = '#F4B084')
writeData(mywb, sheet = "CRM STATEMENT", x = temp)

temp <- read_excel(path = inputfiles[1], sheet = 1)
addWorksheet(mywb, sheet = "CRM REPORT", tabColour = '#F4B084')
writeData(mywb, sheet = "CRM REPORT", x = temp)

temp <- read_excel(path = inputfiles[3], sheet = 1)
addWorksheet(mywb, sheet = "CRM DOCUMENT", tabColour = '#F4B084')
writeData(mywb, sheet = "CRM DOCUMENT", x = documentdata)

addWorksheet(wb = mywb, sheetName = "MO Summary", tabColour = '#F4B084')
addWorksheet(wb = mywb, sheetName = "MO STATEMENT", tabColour = '#F4B084')
addWorksheet(wb = mywb, sheetName = "MO DOCUMENT")
addWorksheet(wb = mywb, sheetName = "Consumption Checks")
addWorksheet(wb = mywb, sheetName = "En & Imp Split")
addWorksheet(wb = mywb, sheetName = "SHADOW SETTLED")




# SUMMARY 
CRMS <- "CRM Summary"

writeData(mywb, CRMS, c(settype3, "CCC", "CSOCDIFFP", "", "Total"), xy = c(1,1))

writeData(mywb, CRMS, "CCC", xy = c(1,8))
writeData(mywb, CRMS, "CSOCDIFFP", xy = c(9,8))

headings <- c("HH", "NHH", "PPA Gen", "Total", "Statement Total", "Check")
for (i in 1:length(headings)) {
  writeData(mywb, CRMS, headings[i], startRow = 9, startCol = i+1)
  writeData(mywb, CRMS, headings[i], startRow = 9, startCol = i+9)
}


dates <- seq(from = as.Date(startdate, format = "%Y-%m-%d"), by = 1, length.out = days_in_month(startdate))
dates <- format(dates, "%d/%m/%Y")

writeData(mywb, CRMS, dates, startRow = 10, startCol = 1)
writeData(mywb, CRMS, dates, startRow = 10, startCol = 9)
writeData(mywb, CRMS, dates, startRow = 11, startCol = 24)

for (i in 1:length(dates)) {
  writeData(mywb, CRMS, dates[i], startRow = 1, startCol = i+1)
}

writeData(mywb, CRMS, rep("E", times = (length(dates)+1)), startRow = 11, startCol = 25)
writeData(mywb, CRMS, rep("P", times = (length(dates)+1)), startRow = 11, startCol = 26)

headings <- c("TRADE DATE", "FROM CURRENCY", 	"TO CURRENCY", "EXCHANGE RATE",	"CMS TIME STAMP")

for (i in 1:length(headings)) {
  writeData(mywb, CRMS, headings[i], startRow = 10, startCol = i+23)
}

writeData(mywb, CRMS, rep(0.8745, time = 32), startRow = 11, startCol = 27)

myletters <- c(LETTERS, sapply(LETTERS, function(x) paste0("A", x)))

for (i in 1:length(dates)) {
  writeFormula(mywb, CRMS, paste0("=SUMIFS(INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"charge_amount\",'CRM DOCUMENT'!$A$1:$BR$1,0)),INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"statement_date\",'CRM DOCUMENT'!$A$1:$BR$1,0)),'CRM Summary'!$", myletters[i+1], "1,INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"charge_name14\",'CRM DOCUMENT'!$A$1:$BR$1,0)),'CRM Summary'!$A2)"), startRow = 2, startCol = i+1)
  
  writeFormula(mywb, CRMS, paste0("=SUMIFS(INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"charge_amount\",'CRM DOCUMENT'!$A$1:$BR$1,0)),INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"statement_date\",'CRM DOCUMENT'!$A$1:$BR$1,0)),'CRM Summary'!", myletters[i+1], "1,INDEX('CRM DOCUMENT'!$A:$BX,0,MATCH(\"charge_name14\",'CRM DOCUMENT'!$A$1:$BR$1,0)),'CRM Summary'!$A3)"), startRow = 3, startCol = i+1)
  
  writeFormula(mywb, CRMS, paste0("=", myletters[i+1], "2+", myletters[i+1], "3"), startRow = 5, startCol = i+1)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!F:F)"), startRow = i+9, startCol = 2)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!G:G)"), startRow = i+9, startCol = 3)
  writeFormula(mywb, CRMS, paste0("=-SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!H:H)"), startRow = i+9, startCol = 4)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!I:I)"), startRow = i+9, startCol = 5)
  writeFormula(mywb, CRMS, paste0("=SUMIFS('CRM STATEMENT'!L:L,'CRM STATEMENT'!K:K,'CRM Summary'!A", i+9, ",'CRM STATEMENT'!J:J,\"CCC\")"), startRow = i+9, startCol = 6)
  writeFormula(mywb, CRMS, paste0("=E", i+9, "+F", i+9), startRow = i+9, startCol = 7)
  
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!N:N)"), startRow = i+9, startCol = 10)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!O:O)"), startRow = i+9, startCol = 11)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!P:P)"), startRow = i+9, startCol = 12)
  writeFormula(mywb, CRMS, paste0("=SUMIF('En & Imp Split'!$A:$A,'CRM Summary'!$A", i+9, ",'En & Imp Split'!Q:Q)"), startRow = i+9, startCol = 13)
  writeFormula(mywb, CRMS, paste0("=SUMIFS('CRM STATEMENT'!P:P,'CRM STATEMENT'!D:D,'CRM Summary'!A", i+9, ",'CRM STATEMENT'!N:N,\"CSOCDIFFP\")"), startRow = i+9, startCol = 14)
  writeFormula(mywb, CRMS, paste0("=M", i+9, "+N", i+9), startRow = i+9, startCol = 15)
  
}


# MAIN TABLE
writeData(mywb, CRMS, "INITIAL CAPACITY", xy = c(17,7))
writeData(mywb, CRMS, paste0(month, "-", year), xy = c(19,7))
writeData(mywb, CRMS, "Totals", xy = c(18,8))

headings <- c("HH", "NHH", "PPA Gen", "Total", "Check against Doc Total")
for (i in 1:length(headings)) {
  writeData(mywb, CRMS, headings[i], startRow = 9, startCol = i+17)
}

headings <- c("NOMINAL CODE", "200005.2", "200005.3", "200005.4")
for (i in 1:length(headings)) {
  writeData(mywb, CRMS, headings[i], startRow = 10, startCol = i+16)
}

writeData(mywb, CRMS, c("Purchases", "Sales"), xy = c(17, 11))

writeFormula(mywb, CRMS, "=B42", startRow = 11, startCol = 18)
writeFormula(mywb, CRMS, "=C42", startRow = 11, startCol = 19)
writeFormula(mywb, CRMS, "=D42", startRow = 11, startCol = 20)
writeFormula(mywb, CRMS, "=SUM(R11:T11)", startRow = 11, startCol = 21)
writeFormula(mywb, CRMS, "U11+U12+S15", startRow = 11, startCol = 22)

writeFormula(mywb, CRMS, "=J42", startRow = 12, startCol = 18)
writeFormula(mywb, CRMS, "=K42", startRow = 12, startCol = 19)
writeFormula(mywb, CRMS, "=-L42", startRow = 12, startCol = 20)
writeFormula(mywb, CRMS, "=SUM(R12:T12)", startRow = 12, startCol = 21)

writeData(mywb, CRMS, "CRM", xy = c(18,15))
writeData(mywb, CRMS, "Invoice Total", xy = c(19,14))
writeData(mywb, CRMS, "Check", xy = c(20,14))
writeFormula(mywb, CRMS, "=SUM(B5:AF5)",  startRow = 15, startCol = 19)
writeFormula(mywb, CRMS, "U11+U12+S15",  startRow = 15, startCol = 20)

writeFormula(mywb, CRMS, "=SUM(B10:B40)",  startRow = 42, startCol = 2)
writeFormula(mywb, CRMS, "=SUM(C10:C40)",  startRow = 42, startCol = 3)
writeFormula(mywb, CRMS, "=SUM(D10:D40)",  startRow = 42, startCol = 4)
writeFormula(mywb, CRMS, "=SUM(E10:E40)",  startRow = 42, startCol = 5)
writeFormula(mywb, CRMS, "=SUM(F10:F40)",  startRow = 42, startCol = 6)
writeFormula(mywb, CRMS, "=E42+F42",  startRow = 42, startCol = 7)

writeFormula(mywb, CRMS, "=SUM(J10:J40)",  startRow = 42, startCol = 10)
writeFormula(mywb, CRMS, "=SUM(K10:K40)",  startRow = 42, startCol = 11)
writeFormula(mywb, CRMS, "=SUM(L10:L40)",  startRow = 42, startCol = 12)
writeFormula(mywb, CRMS, "=SUM(M10:M40)",  startRow = 42, startCol = 13)
writeFormula(mywb, CRMS, "=SUM(N10:N40)",  startRow = 42, startCol = 14)
writeFormula(mywb, CRMS, "=M42+N42",  startRow = 42, startCol = 15)


# FORMATTING
for (ROW in c(2:5, 10:42)) {
  for (COL in c(2:32)) {
    addStyle(mywb, CRMS, style = sterling, rows = ROW, cols = COL)
  }
}

for (ROW in 10:40) {
  for (COL in c(1,9,24)) {
    addStyle(mywb, CRMS, style = DATE, rows = ROW, cols = COL)
  }
}

for (COL in c(2:7, 10:15, 18:22)) {
  addStyle(mywb, CRMS, bluestyle, rows = 9, cols = COL)
}

for (COL in 17:22) {
  addStyle(mywb, CRMS, yellowStyle, rows = 10, cols = COL)
}

addStyle(mywb, CRMS, createStyle(fontSize = 14, textDecoration = 'bold', border = 'TopLeft', borderStyle = 'thick'), rows = 7, cols = 17)
addStyle(mywb, CRMS, createStyle(fontSize = 14, textDecoration = 'bold', border = 'Top', borderStyle = 'thick'), rows = 7, cols = 18:21)
addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = 'TopRight', borderStyle = 'thick'), rows = 7, cols = 22)

addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = 'Left', borderStyle = 'thick'), rows = c(8,9,13,14), cols = 17)
addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = 'Left', borderStyle = 'thick', fontColour = 'red', fgFill = 'yellow'), rows = 10, cols = 17)
addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', fontColour = 'red', fgFill = 'yellow'), rows = 10, cols = 18:21)
addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = "Left", borderStyle = "thick"), rows = c(11,12), cols = 17)

addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = 'Right', borderStyle = 'thick'), rows = c(12:14), cols = 22)
addStyle(mywb, CRMS, createStyle(textDecoration = 'bold', border = 'Right', borderStyle = 'thick', fontColour = 'red', fgFill = 'yellow'), rows = 10, cols = 22)

addStyle(mywb, CRMS, createStyle(numFmt = "[$£]#,##0.00"), rows = 11, cols = c(18:21))
addStyle(mywb, CRMS, createStyle(numFmt = "[$£]#,##0.00"), rows = 12, cols = c(18:21))

addStyle(mywb, CRMS, createStyle(border = 'BottomLeft', borderStyle = 'thick'), rows = 15, cols = 17)
addStyle(mywb, CRMS, createStyle(numFmt = "[$£]#,##0.00", border = 'Bottom', borderStyle = 'thick'), rows = 15, cols = 18:21)
addStyle(mywb, CRMS, createStyle(border = 'BottomRight', borderStyle = 'thick'), rows = 15, cols = 22)

addStyle(mywb, CRMS, createStyle(numFmt = "[$£]#,##0.00", border = 'Right', borderStyle = 'thick'), rows = c(11,8), cols = 22)


# CONDITIONAL FORMATTING
openxlsx::conditionalFormatting(mywb, CRMS, cols = 22, rows = 11, rule = paste0("=OR(V11<-1, V11>1)"), type = 'expression', style = negStyle)
openxlsx::conditionalFormatting(mywb, CRMS, cols = 22, rows = 11, rule = paste0("=AND(V11>=-1, V11<=1)"), type = 'expression', style = posStyle)

openxlsx::conditionalFormatting(mywb, CRMS, cols = 20, rows = 15, rule = paste0("=OR(T15<-1, T15>1)"), type = 'expression', style = negStyle)
openxlsx::conditionalFormatting(mywb, CRMS, cols = 20, rows = 15, rule = paste0("=AND(T15>=-1, T15<=1)"), type = 'expression', style = posStyle)

for (ROW in 10:40) {
  openxlsx::conditionalFormatting(mywb, CRMS, cols = 7, rows = ROW, rule = paste0("=OR(G", ROW, "<-1, G", ROW, ">1)"), type = 'expression', style = negStyle)
  openxlsx::conditionalFormatting(mywb, CRMS, cols = 7, rows = ROW, rule = paste0("=AND(G", ROW, "<=1, G", ROW, ">=-1)"), type = 'expression', style = posStyle)
  openxlsx::conditionalFormatting(mywb, CRMS, cols = 15, rows = ROW, rule = paste0("=OR(O", ROW, "<-1, O", ROW, ">1)"), type = 'expression', style = negStyle)
  openxlsx::conditionalFormatting(mywb, CRMS, cols = 15, rows = ROW, rule = paste0("=AND(O", ROW, "<=1, O", ROW, ">=-1)"), type = 'expression', style = posStyle)
}






# SHADOW SETTLED

headings <- c("Date", "Time P.E", "Date/Time", "Settlement Date", "CCC", "CSOCDIFFP", "CVMO")
for (i in 1:length(headings)) {
  writeData(mywb, SS, headings[i], startRow = 1, startCol = i)
}

temp <- c()
for (i in 2:49) {
  temp <- c(temp, paste0("=DATE(LEFT('CRM REPORT'!L", i, ",4),MID('CRM REPORT'!L", i, ",6,2),MID('CRM REPORT'!L", i, ",9,2))"))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 1)

temp <- c()
for (i in 2:1444) {
  temp <- c(temp, paste0("=A", i, "+1"))
}
writeFormula(mywb, SS, temp, startRow = 50, startCol = 1)

temp <- c()
for (i in 2:49) {
  temp <- c(temp, paste0("=TIME(MID('CRM REPORT'!L", i, ",12,2),MID('CRM REPORT'!L", i, ",15,2),MID('CRM REPORT'!L", i, ",18,2))"))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 2)

temp <- c()
for (i in 2:1444) {
  temp <- c(temp, paste0("=B", i))
}
writeFormula(mywb, SS, temp, startRow = 50, startCol = 2)

temp <- c()
for (i in 2:49) {
  temp <- c(temp, paste0("='CRM REPORT'!L", i))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 3)

temp <- c()
for (i in 50:1492) {
  temp <- c(temp, paste0("=TEXT(A", i, ",\"yyyy-mm-dd\")&\"T\"&TEXT(B", i, ",\"hh:mm:ss\")&\"+00:00\""))
}
writeFormula(mywb, SS, temp, startRow = 50, startCol = 3)

temp <- c()
for (i in 2:49) {
  temp <- c(temp, paste0("='CRM REPORT'!D", i))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 4)

temp <- c()
for (i in 2:1444) {
  temp <- c(temp, paste0("=D", i, "+1"))
}
writeFormula(mywb, SS, temp, startRow = 50, startCol = 4)

temp <- c()
for (i in 2:1492) {
  temp <- c(temp, paste0("=SUMIFS('Consumption Checks'!F:F,'Consumption Checks'!C:C,'SHADOW SETTLED'!C", i, ")*$U$2*SUMIFS('CRM REPORT'!M:M,'CRM REPORT'!L:L,'SHADOW SETTLED'!C", i, ",'CRM REPORT'!J:J,\"FQMCC\")"))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 5)

temp <- c()
for (i in 2:1492) {
  temp <- c(temp, paste0("=E", i, "*$U$3"))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 6)

temp <- c()
for (i in 2:1492) {
  temp <- c(temp, paste0("=SUMIFS('Consumption Checks'!F:F,'Consumption Checks'!C:C,'SHADOW SETTLED'!C", i, ")*$U$4"))
}
writeFormula(mywb, SS, temp, startRow = 2, startCol = 7)

year <- paste0("20",year)

when <- paste0(20, substr(year,3,4), "/", substr(as.numeric(substr(year,3,4))+1,1,2))
then <- paste0(20, substr(as.numeric(substr(as.character(as.numeric(year)-1), 3, 4)),1,2), "/", substr(as.numeric(substr(as.character(as.numeric(year)), 3, 4)),1,2))
then2 <- paste0(20, substr(as.numeric(substr(as.character(as.numeric(year)-2), 3, 4)),1,2), "/", substr(as.numeric(substr(as.character(as.numeric(year)-1), 3, 4)),1,2))
then3 <- paste0(20, substr(as.numeric(substr(as.character(as.numeric(year)-3), 3, 4)),1,2), "/", substr(as.numeric(substr(as.character(as.numeric(year)-2), 3, 4)),1,2))
then4 <- paste0(20, substr(as.numeric(substr(as.character(as.numeric(year)-4), 3, 4)),1,2), "/", substr(as.numeric(substr(as.character(as.numeric(year)-3), 3, 4)),1,2))
then5 <- paste0(20, substr(as.numeric(substr(as.character(as.numeric(year)-5), 3, 4)),1,2), "/", substr(as.numeric(substr(as.character(as.numeric(year)-4), 3, 4)),1,2))

year <- substr(year, 3, 4)

writeData(mywb, SS, c("PCCSUP", "FSOCDIFFP", "PVMO"), startRow = 2, startCol = 15)

temp <- c(11.52, 1.76, 0.029); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 16, x = temp)                                     # these may need amended
temp <- c(21.85, -0.67, 0.029); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 17, x = temp)                                    # review figures with NI
temp <- c(9.19, 0.70, 0); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 18, x = temp)
temp <- c(8.96, 1.61, 0.015); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 19, x = temp)
temp <- c(10.4, 1.25, 0.015); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 20, x = temp)
temp <- c(5.22, 1.3, 0.015); writeData(wb = mywb, sheet = SS, startRow = 2, startCol = 21, x = temp)

if (month.as.number < 10) {
  headings <- c(then, then2, then3, then4, then5)
} else {
  headings <- c(when, then, then2, then3, then4, then5)
}

for (i in 1:length(headings)) {
  writeData(mywb, SS, headings[i], startRow = 1, startCol = i+15)
}

# # FORMATTING
# for (ROW in 1:1492) {
#   for (COL in c(1,4)) {
#     addStyle(mywb, SS, DATE, rows = ROW, cols = COL)
#   }
#   addStyle(mywb, CRMS, TIME, rows = ROW, cols = 2)
#   for (COL in 5:7) {
#     addStyle(mywb, SS, sterling, rows = ROW, cols = 2)
#   }
# }





# EIS

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!A", ROW))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 1)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=A", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 2)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='SHADOW SETTLED'!C", ROW))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 3)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!D", ROW))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 4)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!E", ROW))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 5)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=('Consumption Checks'!P", ROW, "*SUMIFS('CRM REPORT'!$M:$M,'CRM REPORT'!$L:$L,'En & Imp Split'!$C", ROW+2, ",'CRM REPORT'!$J:$J,\"FQMCC\")*'SHADOW SETTLED'!$U$2)*0.8745"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 6)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=('Consumption Checks'!Q", ROW, "*SUMIFS('CRM REPORT'!$M:$M,'CRM REPORT'!$L:$L,'En & Imp Split'!$C", ROW+2, ",'CRM REPORT'!$J:$J,\"FQMCC\")*'SHADOW SETTLED'!$U$2)*0.8745"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 7)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=('Consumption Checks'!R", ROW, "*SUMIFS('CRM REPORT'!$M:$M,'CRM REPORT'!$L:$L,'En & Imp Split'!$C", ROW+2, ",'CRM REPORT'!$J:$J,\"FQMCC\")*'SHADOW SETTLED'!$U$2)*0.8745"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 8)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=F", ROW+2, "+G", ROW+2, "-H", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 9)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=A", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 11)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=D", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 12)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=E", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 13)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=F", ROW+2, "*'SHADOW SETTLED'!$U$3"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 14)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=G", ROW+2, "*'SHADOW SETTLED'!$U$3"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 15)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=H", ROW+2, "*'SHADOW SETTLED'!$U$3"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 16)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=N", ROW+2, "+O", ROW+2, "-P", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 17)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=A", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 19)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=D", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 20)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=E", ROW+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 21)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!P", index, "*'SHADOW SETTLED'!$U$4*VLOOKUP($A", index+2, ",'MO Summary'!$P$8:$S$38,4,0)"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 22)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!Q", index, "*'SHADOW SETTLED'!$U$4*VLOOKUP($A", index+2, ",'MO Summary'!$P$8:$S$38,4,0)"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 23)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("='Consumption Checks'!R", index, "*'SHADOW SETTLED'!$U$4*VLOOKUP($A", index+2, ",'MO Summary'!$P$8:$S$38,4,0)"))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 24)

temp <- c()
for (ROW in 2:1489) {
  temp <- c(temp, paste0("=V", index+2, "+W", index+2, "-X", index+2))
}
writeFormula(mywb, EIS, temp, startRow = 4, startCol = 25)


writeData(mywb, EIS, "CCC", xy = c(1,1))
writeData(mywb, EIS, "CSOCDIFFP", xy = c(11,1))
writeData(mywb, EIS, "CVMO", xy = c(19,1))

headings <- c("Settlement Date", "Trading Date", rep(x = "", times = 3), "HH", "NHH", "PPA Gen", "Total", rep(x = "", times = 4), "HH", "NHH", "PPA Gen", "Total", rep(x = "", times = 4), "HH", "NHH", "PPA Gen", "Total")
for (i in 1:length(headings1)) {
  writeData(mywb, EIS, headings[i], startCol = i, startRow = 3)
}

# # FORMATTING
# for (ROW in 4:1492) {
#   for (COL in c(1,2,11,19)) {
#     addStyle(mywb, EIS, DATE, rows = ROW, cols = COL)
#   }
#   for (COL in c(6:9, 14:17, 22:25)) {
#     addStyle(mywb, EIS, sterling, rows = ROW, cols = COL)
#   }
# }






# CONSUMPTION CHECKS

headings <- c("Settlement Date", "Time P.E", "Date/Time", rep(x = "", times = 2), "MM 596", rep(x = "", times = 8), "TUOS", "MM595", "MM591", "MM598", "MM596")

temp <- c()
for (i in 1:length(headings)) {
  writeData(mywb, CC, headings[i], startRow = 1, startCol = i)
}

writeFormula(mywb, CC, "='SHADOW SETTLED'!A4", startRow = 2, startCol = 1)

temp <- c()
for (ROW in 2:48) {
  temp <- c(temp, paste0("=A", ROW))
}
writeFormula(mywb, CC, temp, startRow = 3, startCol = 1)

temp <- c()
for (ROW in 2:1345) {
  temp <- c(temp, paste0("=A", ROW, "+1"))
}
writeFormula(mywb, CC, temp, startRow = 50, startCol = 1)

temp <- c()
for (ROW in 4:51) {
  temp <- c(temp, paste0("='SHADOW SETTLED'!B", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 2)

temp <- c()
for (ROW in 2:1345) {
  temp <- c(temp, paste0("=B", ROW))
}
writeFormula(mywb, CC, temp, startRow = 50, startCol = 2)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("='SHADOW SETTLED'!C", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 3)

temp <- c()
temp <- seq(from = 3, by = 1, length.out = 48)
writeData(mywb, CC, temp, startRow = 2, startCol = 4)

temp <- c()
for (ROW in 2:1345) {
  temp <- c(temp, paste0("=D", ROW))
}
writeFormula(mywb, CC, temp, startRow = 50, startCol = 4)

time_vector <- c("00:00", "00:30", "01:00", "01:30", "02:00", "02:30", "03:00", "03:30", "04:00", "04:30", "05:00", "05:30",
                 "06:00", "06:30", "07:00", "07:30", "08:00", "08:30", "09:00", "09:30", "10:00", "10:30", "11:00", "11:30",
                 "12:00", "12:30", "13:00", "13:30", "14:00", "14:30", "15:00", "15:30", "16:00", "16:30", "17:00", "17:30",
                 "18:00", "18:30", "19:00", "19:30", "20:00", "20:30", "21:00", "21:30", "22:00", "22:30", "23:00", "23:30")

writeData(mywb, CC, time_vector, startRow = 2, startCol = 5)

temp <- c()
for (ROW in 2:1345) {
  temp <- c(temp, paste0("=E", ROW))
}
writeFormula(mywb, CC, temp, startRow = 50, startCol = 5)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=VLOOKUP(A", ROW, ",'", accdrive, ":/GeneralAccounts/Settlement/NIE MM/D+4 Initial 20 MM/[NI MASTER MM Messages D+4.xlsx]MM596'!$A:$AY,D", ROW, ",0)"))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 6)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=A", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 12)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=D", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 13)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=E", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 14)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=VLOOKUP(L", ROW, ",'", accdrive, ":/GeneralAccounts/Settlement/PT_500057 SEMO Shadow Settlement/TUoS, CAIR, SSS/[SONI Bill Summary.xlsx]2023.24 CALENDAR'!$A:$AX,M", ROW, ",0)"))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 15)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=VLOOKUP($L", ROW, ",'", accdrive, ":/GeneralAccounts/Settlement/NIE MM/D+4 Initial 20 MM/[NI MASTER MM Messages D+4.xlsx]MM595 LA'!$A:$AZ,$M", ROW, ",0)/1000"))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 16)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=VLOOKUP($L", ROW, ",'", accdrive, ":/GeneralAccounts/Settlement/NIE MM/D+4 Initial 20 MM/[NI MASTER MM Messages D+4.xlsx]MM591 LA'!$A:$AY,$M", ROW, ",0)/1000"))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 17)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=VLOOKUP($L", ROW, ",'", accdrive, ":/GeneralAccounts/Settlement/NIE MM/D+4 Initial 20 MM/[NI MASTER MM Messages D+4.xlsx]MM598'!$A:$AZ,$M", ROW, ",0)/1000"))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 18)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=F", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 19)

temp <- c()
for (ROW in 2:1393) {
  temp <- c(temp, paste0("=Q", ROW, "+P", ROW, "-R", ROW, "+S", ROW))
}
writeFormula(mywb, CC, temp, startRow = 2, startCol = 20)

totals <- seq(from = 49, by = 48, length.out = days_in_month(startdate))

temp <- c()
for (total in totals) {
  temp <- c(temp, paste0("=IFERROR(MIN(MAX(SUMIFS('CRM REPORT'!M:M,'CRM REPORT'!D:D,'SHADOW SETTLED'!D", index, ",'CRM REPORT'!J:J,\"CBSOC\"),0),-SUMIFS('CRM REPORT'!S:S,'CRM REPORT'!D:D,'SHADOW SETTLED'!D", index, ",'CRM REPORT'!P:P,\"CSHORTDIFFPTRACK\"))*(SUMIFS('CRM REPORT'!S:S,'CRM REPORT'!D:D,'SHADOW SETTLED'!D", index, ",'CRM REPORT'!P:P,\"CSHORTDIFFPTRACK\")/SUMIFS('CRM REPORT'!S:S,'CRM REPORT'!D:D,'SHADOW SETTLED'!D", index, ",'CRM REPORT'!P:P,\"CSHORTDIFFPTRACK\")),0)"))
}
for (i in 1:7) {
  writeFormula(wb = mywb, sheet = CC, x = temp[i], startRow = totals[i], startCol = 22)
}

temp <- c()
for (total in totals) {
  temp <- c(temp, paste0("=IF(AND(SUM(U", total-47, ":U", total, ")<0.01,SUM(U", total-47, ":U", total, ")>-0.01),\"TRUE\",\"FALSE\")"))
}
for (i in 1:7) {
  writeFormula(wb = mywb, sheet = CC, x = temp[i], startRow = totals[i], startCol = 23)
}

# FORMATTING
for (ROW in c(2:1393)) {
  for (COL in c(1, 12)) {
    addStyle(mywb, CC, DATE, rows = ROW, cols = COL)
  }
  addStyle(mywb, CC, TIME, rows = ROW, cols = 2)
  for (COL in c(6,16:20)) {
    addStyle(mywb, CC, NUMBER, rows = ROW, cols = COL)
  }
}

saveWorkbook(mywb, outputfile, TRUE)

# add to Invoice Backup

ivbuoutput <- paste0(accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Energy + Imp/Invoice Backups/", documentidnumber, " BACKUP - ", initials, ".xlsx")

# sales
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!R12"), startRow = 23, startCol = 9)
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!S12"), startRow = 23, startCol = 10)
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!T12"), startRow = 23, startCol = 11)

# purchases
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!R11"), startRow = 45, startCol = 9)
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!S11"), startRow = 45, startCol = 10)
writeFormula(ivbu, BU, paste0("='", accdrive, ":/GeneralAccounts/Settlement/", PTunit, " SEMO Shadow Settlement/Capacity/", settype2, "/[20", year, " ", month.as.number,  ". ", PTunit, " SEMO ", month, " ", year, " ", settype, " CAPACITY Shadow Settlement - ", initials, ".xlsx]CRM Summary'!T11"), startRow = 45, startCol = 11)


saveWorkbook(ivbu, ivbuoutput, TRUE)
