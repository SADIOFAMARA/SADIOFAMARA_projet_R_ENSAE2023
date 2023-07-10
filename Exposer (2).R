#1. Quelques packages et fonctions necessaires.

library(readxl)   # Ce package permet d'importer des données à partir de fichiers Excel
# Exemple:  data <- read_excel("chemin/vers/votre/fichier.xlsx", sheet = "nom_de_la_feuille")


library(writexl)  # Ce package permet d'exporter des données depuis R vers des fichiers Excel (.xlsx) sans nécessiter Microsoft Excel lui-même. Il offre une interface simple pour créer des fichiers Excel avec différentes feuilles de calcul et formats.
library(openxlsx) # Ce package permet de lire, d'écrire et de manipuler des fichiers Excel (.xlsx). Il offre des fonctionnalités avancées telles que la création de feuilles de calcul, l'ajout de graphiques, la gestion des styles et des formats de cellules, etc
library(xlsx)     # Ce package fournit des fonctions pour lire, écrire et manipuler des fichiers Excel (.xlsx) en R. Il permet également de créer des feuilles de calcul, de gérer les formats de cellules et de fusionner des cellules
library(r2excel)  #

read_excel() (package readxl) : Cette fonction permet de lire un fichier Excel (.xlsx ou .xls) et de charger les données dans R en tant que data frame.

# les fonctions les plus importantes

write.xlsx() (package openxlsx) : #Cette fonction permet d'écrire un data frame ou une liste dans un fichier Excel (.xlsx).


writeData() #(package writexl) : Cette fonction permet d'écrire des données dans un fichier Excel en utilisant la librairie writexl.

getSheetNames() #(package openxlsx) : Cette fonction permet d'obtenir les noms de toutes les feuilles de calcul présentes dans un fichier Excel.

# Installer le package r2excel
# install.packages("r2excel")

# Charger le package r2excel


# Fonction pour écrire plusieurs data frames dans un fichier Excel avec chaque data frame dans une feuille de calcul séparée
# write_xlsx_from_dataframes(list_of_dataframes, file_path)

# Fonction pour écrire plusieurs listes d'objets dans un fichier Excel avec chaque liste dans une feuille de calcul distincte
# write_xlsx_from_lists(list_of_lists, file_path)

# Fonction pour écrire un workbook (fichier Excel) existant dans un fichier Excel (.xlsx)
# write_xlsx_from_workbook(workbook, file_path)

# Fonction pour ajouter une nouvelle feuille de calcul à un fichier Excel existant
# add_sheet(workbook, sheet_name)

# Fonction pour ajouter un tableau formaté à une feuille de calcul Excel existante
# add_table(workbook, sheet_name, table_data, start_row, start_col)



# 1.Manipulation elementaires de excel via le logiciel R

# Partie 2

library(openxlsx)

# Créer un nouveau fichier Excel

wb <- createWorkbook()

# Créer une nouvelle feuille de calcul dans le fichier
addWorksheet(wb, "MaFeuille")

# Ajouter des données à la feuille de calcul
data <- data.frame(A = 1:5, B = letters[1:5])
writeData(wb, "MaFeuille", data, startCol = 1, startRow = 1)

# Sauvegarder le fichier Excel
saveWorkbook(wb, "mon_fichier.xlsx")

#Lecture de fichiers Excel dans le langage de programmation R

Data1 < - read_excel("Sample_data1.xlsx")
Data2 < - read_excel("Sample_data2.xlsx")

# Modification de fichiers

Data1$Pclass <- 0
Data2$Embarked <- "S"

# Suppression de contenu de fichiers
#La variable ou l’attribut est supprimé des jeux de données Data1 et Data2 contenant des fichiers Sample_data1.xlsx et Sample_data2.xlsx.

# Deleting from files
Data1 <- Data1[-2]

Data2 <- Data2[-3]

# Merging Files
Data3 <- merge(Data1, Data2, all.x = TRUE, all.y = TRUE)
head(Data3)

# Création d'un élément dans le jeu de données Data2
Data1$Num < - 0

# Creating feature in Data2 dataset
Data2$Code < - "Mission"

# Printing the data
head(Data1)
head(Data2)

# Comment écrire plusieurs fichiers Excel à partir de valeurs de colonne - Programmation R
# invoking the required packages
library(xlsx)


# Fallou Badji

filename<-"travail.xlsx"
donnee <- structure(list(Annee = c(2013, 2014, 2015, 2016, 2017),
                         Nombre_de_saillies = c(15,18, 32, 33, 25)), class = "data.frame", row.names = c(NA, -5L
                         ))
donnee<-head(iris)
wb <- createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName = "addDataFrame1")
addDataFrame(donnee, sheet, col.names = TRUE, row.names = TRUE)

#création des styles
cs1 = CellStyle(wb) +
   Font(wb, isBold=TRUE) +
   Fill(backgroundColor="lavender", foregroundColor="lavender",
        pattern="SOLID_FOREGROUND") +
   Alignment(h="ALIGN_CENTER",v="VERTICAL_CENTER")+Border(color = "black", position = c("TOP","BOTTOM","LEFT","RIGHT"))
cs2 = CellStyle(wb) +
   Font(wb, isBold=TRUE) +
   Fill(backgroundColor="#eb9853", foregroundColor="#eb8953",
        pattern="SOLID_FOREGROUND") +
   Alignment(h="ALIGN_CENTER",v="VERTICAL_CENTER")+Border(color = "black", position = c("TOP","BOTTOM","LEFT","RIGHT"))

rows <- getRows(sheet)
cells <- getCells(rows)
#on crée un df avec les références des lignes (X1) et des colonnes (X2) ce qui simplifie la selection des cellules
ref_cell<-data.frame(t(do.call("cbind", strsplit(names(cells),"\\."))),stringsAsFactors = F)
l_ind<-which(ref_cell$X1 ==1) #recherche les cellules de la ligne 1
lapply(l_ind, function(i) setCellStyle(cells[[i]], cs1)) #on applique le style cs1 sur les cellules de la ligne 1

l_ind<-which(ref_cell$X1  %in% 3:5) #recherche des cellules des ligne 2 et 3
lapply(l_ind, function(i) setCellStyle(cells[[i]], cs2)) #on applique le style cs2 sur les cellules des ligne 2 et 3

saveWorkbook(wb,"travail.xlsx")
xlsx.openFile("travail.xlsx")


#creating a data frame
data_frame <- data.frame(col1=c(1:10),
                         col2=c("Anna","Mindy","Bindy",
                                "Tindy","Ron",
                                "Charles","Zoe","Dan",
                                "Lincoln","Burrows"),
                         col3=c("CS","CA","Eco","Eco","CA",
                                "Eco","CS","CS","CS","Eco"))

# segregating data based on the col3 values
data_mod <- split(data_frame, data_frame$col3)

# printing the obtained groups
print("Segregated dataframes")
print(data_mod)

# obtenier des  differents dimensions
# nombre 
size <- length(data_mod)

#creer les nombres de listes equivalents
# to the size of the generated groups
lapply(1:size,
       function(i)
          write.xlsx(data_mod[[i]],
                     file = paste0("/Users/mallikagupta/Desktop/",
                                   names(data_mod[i]), ".xlsx")))

# invoking the required packages
library("xlsx")
library("dplyr")

# creating a data frame
data_frame <- data.frame(col1 = c(1:10),
                         col2=c("Anna","Mindy","Bindy",
                                "Tindy","Ron",
                                "Charles","Zoe","Dan",
                                "Lincoln","Burrows"),
                         col3=c("CS","CA","Eco","Eco","CA",
                                "Eco","CS","CS","CS","Eco"),
                         col4=c(1,3,2,2,3,4,1,4,1,2))

# segregating data based on the boolean condition of
# whether the col3 values is greater that 2 or not
data_mod <- data_frame %>%
   group_split(col4>2)

# printing the different groups created
print("Segregated data frames")

# two groups are created
print(data_mod)


Examples

# 2. manipulation de excel via R avec r2excel

library(xlsx)
library(r2excel)
# Create an Excel workbook. Both .xls and .xlsx file formats can be used.
filename<-"-example.xlsx"
wb <- createWorkbook(type="xlsx")
# Create a sheet in that workbook
sheet <- xlsx::createSheet(wb, sheetName = "example1")

# Add header
#+++++++++++++++++++++++++++++++
# Create the Sheet title and subtitle
xlsx.addHeader(wb, sheet, value="Excel file written with r2excel packages",
               level=1, color="darkblue", underline=2)         
xlsx.addLineBreak(sheet, 2)

# Add paragraph : Author
#+++++++++++++++++++++++++++++++
xlsx.addParagraph(wb, sheet, value="Author : Alboukadel KASSAMBARA. \n@:alboukadel.kassambara@gmail.com.\n Website : http://ww.sthda.com", isItalic=TRUE, colSpan=5, rowSpan=4, fontColor="darkgray", fontSize=14)
xlsx.addLineBreak(sheet, 3)

getwd()
# Add table
#+++++++++++++++++++++++++++++
# add iris data using default settings
library(readxl)
data("iris")
Data_MCA <- read_excel("Data_MCA.xlsx")
View(Data_MCA)
xlsx.addHeader(wb, sheet, value="Add iris table using default settings", level=2)
xlsx.addLineBreak(sheet, 1)
xlsx.addTable(wb, sheet, 
              
              
              iris %>%
                 labelled::set_variable_labels(
                    Petal.Length = "Longueur du pétale",
                    Petal.Width = "Largeur du pétale"
                 ) %>%
                 tbl_summary(label = Species ~ "Espèce") %>%
                 add_n(
                    statistic = "{n}/{N}",
                    col_label = "**Effectifs** (observés / total)",
                    last = TRUE,
                    footnote = TRUE
                 )
              
              , startCol=2)
xlsx.addLineBreak(sheet, 2)

# Customized table
xlsx.addHeader(wb, sheet, value="Customized table", level=2)
xlsx.addLineBreak(sheet, 1)
xlsx.addTable(wb, sheet, data= head(iris),
              fontColor="darkblue", fontSize=14,
              rowFill=c("white", "lightblue")
)
xlsx.addLineBreak(sheet, 2)


# Add paragraph
#+++++++++++++++++++++++++++++
xlsx.addHeader(wb, sheet, "Add paragraph", level=2)
xlsx.addLineBreak(sheet, 2)
paragraph="Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged."
xlsx.addParagraph(wb, sheet, paragraph, fontSize=14, isItalic=TRUE, 
                  fontColor="darkred", backGroundColor="gray")
xlsx.addLineBreak(sheet, 2)


# Add box plot
#+++++++++++++++++++++++++++++
data(iris)
xlsx.addHeader(wb, sheet, " Add Plot", level=2)
xlsx.addLineBreak(sheet, 1)
plotFunction<-function(){boxplot(len ~ dose, data = iris, col = 1:3)}
xlsx.addPlot(wb, sheet, plotFunction())

# save the workbook to an Excel file and write the file to disk.
xlsx::saveWorkbook(wb, filename)

#open file
xlsx.openFile(filename)

ft <- flextable::flextable(head(mtcars))
ft
# color some cells in blue
ft <- flextable::bg(ft, i=ft$body$dataset$disp>200, j=3, bg = "#7ed6df", part = "body")
# color a few cells in yellow
ft <- flextable::bg(ft, i=ft$body$dataset$vs==0, j=8, bg = "#FCEC20", part = "body")
# export your flextable as a .xlsx in the current working directory
exportxlsx(ft, path = "X:/temp_del/excel_file.xlsx")



#Exemple pratique exportation de tableaux redalise a partir d'une base vers excel

install.packages("expss")
library(expss)
library(openxlsx)
# Sauvegarder la base de données mtcars dans un fichier CSV
write.xls(mtcars, file = "C/Users/PATE DIAGNE/Desktop/add/Excel/mtcars.csv", row.names = FALSE)

str(mtcars)
tail(mtcars)

names(mtcars)
mtcars = apply_labels(mtcars,
                      mpg = "Miles/(US) gallon",
                      cyl = "Number of cylinders",
                      disp = "Displacement (cu.in.)",
                      hp = "Gross horsepower",
                      drat = "Rear axle ratio",
                      wt = "Weight (lb/1000)",
                      qsec = "1/4 mile time",
                      vs = "Engine",
                      vs = c("V-engine" = 0,
                             "Straight engine" = 1),
                      am = "Transmission",
                      am = c("Automatic" = 0,
                             "Manual"=1),
                      gear = "Number of forward gears",
                      carb = "Number of carburetors"
)

mtcars_table = mtcars %>% 
   cross_cpct(
      cell_vars = list(cyl, gear),
      col_vars = list(total(), am, vs)
   ) %>% 
   set_caption("Table 1")

mtcars_table
wb = createWorkbook()
sh = addWorksheet(wb, "Tables")
xl_write(mtcars_table, wb, sh)
saveWorkbook(wb, "table1.xlsx", overwrite = TRUE)
banner = with(mtcars, list(total(), am, vs))

list_of_tables = lapply(mtcars, function(variable) {
   if(length(unique(variable))<7){
      cro_cpct(variable, banner) %>% significance_cpct()
   } else {
      # if number of unique values greater than seven we calculate mean
      cro_mean_sd_n(variable, banner) %>% significance_means()
      
   }
   
})

wb = createWorkbook()
sh = addWorksheet(wb, "Tables")

xl_write(list_of_tables, wb, sh, 
         # remove '#' sign from totals 
         col_symbols_to_remove = "#",
         row_symbols_to_remove = "#",
         # format total column as bold
         other_col_labels_formats = list("#" = createStyle(textDecoration = "bold")),
         other_cols_formats = list("#" = createStyle(textDecoration = "bold")),
)

saveWorkbook(wb, "report.xlsx", overwrite = TRUE)

getwd()
# //////////////////////////////////////////////////



