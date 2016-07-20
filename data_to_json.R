library(RJSONIO)

data <- read.csv('https://docs.google.com/spreadsheets/d/1nlYOuCFByr6Cmy4XA1hlOt69vQil4hVfHeYypsxoll0/pub?output=csv', header=TRUE, stringsAsFactors=FALSE)

data_list <- list()

for (row_i in seq_len(nrow(data))) {
  
  data_list[[row_i]] <- data[row_i, , drop=TRUE]
}

data_json <- toJSON(data_list)

test <- fromJSON(data_json)

writeLines(data_json, 'data.json')
