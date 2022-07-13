#DataRead.R																						
																						
library(readxl)																						
library(cellranger)																						

## Excel file with Name, Tab, Range columns. Blank cells for auto detection of table boundaries.
address <- read_excel("TestTableAddress.xlsx")																						
																						
address[c("StartLetter", "StartNumber", "EndLetter", "EndNumber")] <- NA																						
																						
for (i in 1:nrow(address)) {																						
	address$StartLetter[i] <- letter_to_num(regmatches(address$Range[i], gregexp("[[:alpha:]]+", address$Range[i]))[[1]][1])																					
	address$StartNumber[i] <- regmatches(address$Range[i], gregexpr("[[:digit:]]+", address$Range[i]))[[1]][1]																					
	address$EndLetter[i] <- letter_to_num(regmatches(address$Range[i], gregexp("[[:alpha:]]+", address$Range[i]))[[1]][2])																					
	address$EndNumber[i] <- regmatches(address$Range[i], gregexpr("[[:digit:]]+", address$Range[i]))[[1]][2]																					
}																						
																						
listOfObjs <- list()																						
																						
for (i in 1:nrow(address)) {																						
	df <- read_excel("..", sheet = (address$Tab[i]),
			 cell_limits(c(address$StartNumber[i], address$StartLetter[i]), c(address$EndNumber[i], address$EndLetter[i])),
			 col_names = F)														
	listofObjs[[i]] <- df																					
	names(listOfObjs)[i] <- address$Name[i]																					
	assign(paste(address$Name[i]), df)																					
}
