install.packages("readxl")

library(readxl)

data1 <- read_excel("H:\\Downloads\\free_text.xlsx")#path to your file
basepath="C:\\Users\\c1049033\\Documents\\misc\\"
data1[is.na(data1)] <- ""

###############################################SEPARATE RIS FILES FOR EACH COLUMN
myvars<-c("1.4a","2.4a","3.4a")

for (myvar in myvars){
  mylist<-c()
  
  for(i in 1:nrow(data1)) {
    row <- data1[i,]
    
    mylist<-c(mylist,"TY  - JOUR")
    mylist<-c(mylist,sprintf("ID  - %s",toString(row['gen_ID'])))
    mylist<-c(mylist,sprintf("TI  - %s",toString(row[myvar])))
    mylist<-c(mylist,sprintf("AB  - %s",toString(row[myvar])))
    mylist<-c(mylist,"ER  - ")
    mylist<-c(mylist,"")
  }
  
  filename=sprintf("%s%s.ris",basepath, gsub("\\.", "", myvar))
  file.create(filename)
  lapply(mylist, write, filename, append=TRUE)
  print(sprintf("Created %s", filename))
  
}

##########################################################MERGED

mylist<-c()

for(i in 1:nrow(data1)) {
  row <- data1[i,]
  all_data=""
  
  for (myvar in myvars){
    all_data=sprintf("%s %s", all_data, toString(row[myvar]))
    
  }
  all_data=gsub("  ", " ", all_data)
  all_data=trimws(all_data)
  
  mylist<-c(mylist,"TY  - JOUR")
  mylist<-c(mylist,sprintf("ID  - %s",toString(row['gen_ID'])))
  mylist<-c(mylist,sprintf("TI  - %s", all_data))
  mylist<-c(mylist,sprintf("AB  - %s", all_data))
  mylist<-c(mylist,"ER  - ")
  mylist<-c(mylist,"")
}

filename=sprintf("%s%s.ris",basepath, "all_merged")
file.create(filename)
lapply(mylist, write, filename, append=TRUE)
print(sprintf("Created %s", filename))

