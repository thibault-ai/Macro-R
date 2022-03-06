############################################################################################################
#                      ISSEA 2021 - 2022 IAS3                                                              #
#                                                                                                          #
#        COURS DE LOGICIELS STATISTIQUES : Projet sur R                                                    #
#                                                                                                          #
#		             Projet_R_2022_R                                                                           #
# Membres : TATOU DEKOU THIBAULT, MONSI EMEFA SOPHIE, LOUCOUBAMA PRINCY CODA KENAN                         #        
############################################################################################################

############################################################################################################
###########fonction d'importation des fichiers##############################################################
Import<-function(pathfichier=file.choose()){
  lancerImport=winDialog(type = "yesno","L'importation va nécessiter le package stringr. Voulez vous le charger?")
  if(lancerImport=="YES"){
    require(stringr,quietly = T)##pour l'importation le package stringr sera utilisé###########################################
    extension = str_sub(pathfichier,start=nchar(pathfichier)-3L,end = nchar(pathfichier))##on commence par extraire les quatres derniers caractères du lien vers le fichier pour 
    ##déduire l'extension de celui-ci
    ##nous construisons le vecteur qui va aider à l'identification des extensions
    Extensions=c(".txt",".csv",".xls",".xlsx",".sav",".dta")
    ##on déduit l'extension véritable du fichier en cherchant dans le vecteur précédent l'élément qui contient les caractères extrait au début
    extension_f=str_subset(Extensions,extension)
    
    ##pour lancer l'importation nous lisons la première ligne du fichier
    a=readLines(pathfichier,n=1)
    ##s'il s'agit d'un fichier texte comme le controlé l'instruction ci'après
    if (extension_f==".txt"|extension_f==".csv"){
      ##nous cherchons quel est le délimiteur
      ##nous construisons un dataframe avec les délimiteurs possible et le nombre de fois que ceux ci apparaissent dans la première ligne
      tab=str_count(a,pattern="\t")
      virg=str_count(a,pattern=",")
      pvirg=str_count(a,pattern=";")
      separ=list(separateur=c("\t",",",";"),nombre=c(tab,virg,pvirg))
      separ_frame=as.data.frame(separ)
      ##de la on recherche le délimiteur qui apparait au moins une fois
      index=which(separ_frame[,2]!=0,arr.ind = T)
      #puis l'on va chercher le délimiteur en question dans le dataframe précédent 
      separateur=separ_frame[,1][index]
      ##et enfin on peut importer le fichier
      donnees=as.data.frame(read.table(pathfichier,header = T,sep = separateur,dec="."))
    }
    ##s'il s'agit d'un fichier spss ou stata
    if (extension_f==".sav"|extension_f==".dat"){
      ##nous utilisons le package foreign pour l'importation
      require(foreign,quietly = T)
      if (extension_f==".sav"){
        donnees=as.data.frame(read.spss(pathfichier))
      }
      else{
        donnees=as.data.frame(read.dta(pathfichier))
      }
    }
    ###pour les fichiers excel, nous utilisons le package xlsx
    if (extension_f==".xls"|extension_f==".xlsx"){
      excel=winDialog(type = "yesno","L'importation de votre fichier excel va nécessiter le package xlsx. L'avez vous déja installé?")
      if (excel=="YES"){
        require(xlsx,quietly = T)
        donnees=as.data.frame(read.xlsx2(pathfichier,sheetIndex=1, header =T)) 
      }
    }
    donnees
  }

}
#############################################################################################################
##############fonction pour détecter les types des variables   ##############################################
##cette fonction permet de detecter les deux variables numériques du jeu de données
detectvar<-function(donnees){
  type_var=as.data.frame(list(numero=c(1,2,3),
                              type=c(mode(donnees[,1]),mode(donnees[,2]),mode(donnees[,3]))))
  ##nous construisons un dataframe avec les numéros et les modes des variables telles que fournies
  ##dans le jeu de données
  ##on determine la position de la colonne chaine de caractere
  index=which(type_var[,2]=="character",arr.ind = T)
  #a partir de laquelle est construite la variable catégorielle
  varqual=donnees[,index]
  ##le colonne catégorielle est supprimée du dataframe précedent 
  col_num=type_var[-index,]
  ##et on construit les deux variables numériques à partir des deux numéros qui ne correspondait pas
  varnum1=donnees[,col_num[,1][1]]
  varnum2=donnees[,col_num[,1][2]]
  ###on sauvegarde les numeros des colonnes dans l'ordre voulu
  index_col=c(col_num[,1][1],col_num[,1][2],index)
  noms=str_split(names(donnees),pattern=" ",simplify = T)
  nom_col=c(noms[index_col[1]],noms[index_col[2]],noms[index_col[3]])
  #pour construire la liste des variables avec leur noms
  variables=list(varnum1,varnum2,varqual,nom_col)
  variables
}
#####################################################################################################
###########fonction pour réaliser les statistiques descriptives######################################
##cette fonction a pour but de fournir l'essentiel des statistiques descriptives demandées par le 
##projet
statdesc<-function(donnees){
  ##les variables du jeu de données sont d'abord déterminées
  variables=detectvar(donnees)
  noms=str_split(names(donnees),pattern=" ",simplify = T)
  ##pour faire un résumé de base de chacune d'elle (noms, type et effectif)
  resume=as.data.frame(list(Noms=c(noms[1],noms[2],noms[3]),
                            type=c(mode(donnees[,1]),mode(donnees[,2]),mode(donnees[,3])),
                            Effectif=c(length(donnees[,1]),length(donnees[,2]),length(donnees[,3]))))
  ##le package graphics est utilisé pour les graphiques
  require(graphics,quietly = T)
  #nous construisons d'abord les deux histogrammes des variables numériques
  for (i in 1:2) {
    png(paste("Histogramme",variables[[4]][i],".png"),height = 2000,width =2000,
        res = 250,pointsize = 8)
    hist(variables[[i]],col = "yellow", probability = T,
         main = paste("Histogram of",variables[[4]][i]),xlab =variables[[4]][i])
    lines(density(variables[[i]]),col="blue",lty="solid",lwd=1)
    dev.off()
  }
  #puis les boites à moustaches des deux variables numériques
  for (i in 1:2) {
    png(paste("Boxplot",variables[[4]][i],".png"),height = 2000,width = 2000,
        res = 250,pointsize = 8)
    boxplot(variables[[i]],col = "Blue",xlab =variables[[4]][i],
            main=paste("Boxplot",variables[[4]][i]))
    dev.off()
  }
  ##le tableau des effectifs de la variable catégorielle est construit
  fac=factor(variables[[3]])
  tableau_effect=as.data.frame(list(modalite=c(unique(variables[[3]]),"Total"),
                                    effectif=c(tabulate(fac),sum(tabulate(fac)))))
  ###nous construisons les boites à moustaches selon les classes de la variables catégorielles
  for (i in 1:2) {
    png(paste("Boxplot_multi_cellules",variables[[4]][i],".png"),height = 2000,
        width = 2000,res = 250,pointsize = 8)
    boxplot(variables[[i]]~fac,col=1:24,xlab = variables[[4]][3],ylab = variables[[4]][i],
            main=paste("Boxplot_multi_cellules",variables[[4]][i]))
    dev.off()
  }
  png(paste("Nuage de points de la variable",variables[[4]][1],
            "en fonction de",variables[[4]][2],".png"),
      height = 2000,width = 2000,res = 250,pointsize = 8)
  
  ##le nuage de point entre les deux variables numériques est construit
  plot(variables[[1]]~variables[[2]],col=as.numeric(fac),
       xlab = variables[[4]][2],ylab = variables[[4]][1],
       main=paste("Nuage de points de la variable",variables[[4]][1],
                  "en fonction de",variables[[4]][2]))
  dev.off()
  ##nous calculons le coefficient de corrlations
  coef_corr=cor(variables[[1]],variables[[2]],method="pearson")
  ###nous constrisons un element de type list pour sauvergader les résultats numérique
  statistiques=list(resume,tableau_effect,coef_corr)
  statistiques
}
###################################################################################################################################"
###algorithme de descente du gradient à pas fixe.###############################################################################
#pour l'estimation des coefficients de régression nous avons utilisé l'algorithme de descente du gradient à pas fixe
GradientDescent<-function(vary,varx){
  m=length(vary)
  ones=rep(x = 1,times=m)
  varx=cbind(ones,varx)
  theta=rep(x = 0,times=2)
  iterations=30000;
  alpha=0.01
  a=crossprod(varx,varx)
  for (iter in 1:iterations){
    theta=theta-alpha*(1/m)*(crossprod(a,theta)-crossprod(varx,vary))
  }
  theta
}
####################################################################################################################################
##############fonction pour construire le tableau des coefficients#################################################################
####le tableau des coefficeints de la regression est ensuite construit grace à la fonction ci apres
tableauCoef<-function(vary,varx,er=0.05){
  m=length(vary)
  theta=GradientDescent(vary,varx)  ##les coefficients sont calculés
  residus=vary-(theta[1]+theta[2]*varx)  ##les résidus sont déterminés
  sigma_carre=(m/(m-2))*var(residus) ##estimation de la variance des résidus
  var_t0=sqrt((sigma_carre*sum(varx^2))/(m*m*var(varx))) ##variance de l'estimateur de l'ordonnée à l'origine
  t_0 = theta[1]/var_t0                     ###statistique du test de significativité de l'intercept
  precision_0=qt(1-(er/2),m-2)*var_t0          ###précision sur la l'estimation du paramètre pour la construction de l'intervalle de confiance
  C0_ = t_0 - precision_0
  C0 = t_0 + precision_0
  p_valeur_0 = min(c(2*pt(t_0,m-2),2*(1-pt(t_0,m-2))))        ###calcul de la pvaleur du test de significativité de l'intercept
  var_t1=sqrt(sigma_carre/(m*var(varx)))       ##variance de l'estimation de la pente
  t_1 = theta[2]/var_t1             ##statistique du test de significativite de la pente
  precision_1=qt(1-(er/2),m-2)*var_t1      ###précision sur l'estimation de la pente pour la construction des intervalles de confiance
  C1_ = t_1 - precision_1
  C1 = t_1 + precision_1
  p_valeur_1 = min(c(2*pt(t_1,m-2),2*(1-pt(t_1,m-2))))  ###calcul de la pvaleur du test de significativité de la pente
  ###le dataframe ci dessous regroupe les résultats de l'estimation tels que demandés
  Coefficients= as.data.frame(list(C=c("Intercept","Pente"), v=theta,
                     I=c(paste("[",round(C0_,digits = 4),";",round(C0,digits = 4),"]"),
                     paste("[",round(C1_,digits = 4),";",round(C1,digits = 4),"]")),
                     stat=c(round(t_0,digits = 3),round(t_1,digits = 3)),
                     p=c(p_valeur_0,p_valeur_1)))
  row.names(Coefficients)=c(1,2)
  colnames(Coefficients)=c("Coefficient","Valeur estimée",
                           "Intervalle de confiance à 1-er",
                           "t-statistic","p-value")
  Coefficients
}
##########################################################################################################################################
##################affichage des statistique ###################################################################################################
##cette fonction vise à construire les statistiques servant à la validation du modele
statval<-function(vary,varx,er=0.05){
  m=length(vary)
  theta=GradientDescent(vary,varx) ##a partir des variables les coefficients sont de nouveau estimés
  ypred=theta[1]+theta[2]*varx    ###la valeur prédite de la variable explicative est calculée
  residus=vary-ypred ###puis les résidus
  SCE=sum(residus^2) ###la somme des carrés des résidus
  score=var(ypred)/var(vary) ##le coefficient de détermination du modele de regression
  score_ajuste = 1-(1-score)*((m-1)/(m-2)) ###la valeur ajusté de ce paramètre
  Stat_Fisher=var(ypred)/(var(residus)/(m-2)) ###la statistique du test de fisher
  p_valeur_F = 1-pf(Stat_Fisher,1,m-2)  ##la pvaleur de ce test
  
  require(stats,quietly = T)
  test_kol=suppressMessages(ks.test(residus,"pnorm",mean(residus),sd(residus))) ##test de kolmogorov
  stat_kol=test_kol[[1]]         ###la statistique du test
  p_valeur_kol=test_kol[[2]] ##la pvaleur du test de kolmogorov 
  test_sh=shapiro.test(residus)         ####le test de shapiro wilk
  stat_sh=test_sh[[1]]              ###la statistique de ce test 
  p_valeur_sh=test_sh[[2]]              ###sa pvaleur
  AIC=log(SCE/m) + (2/m)         ###le critère d'information d'AKAIKE
  ####cela est consigné dans le dataframe suivant
  tableau_stat = as.data.frame(list(Statistique=c("Somme des carrés des résidus",
                                                  "Score","Score ajusté","F-statistique",
                                                  "P-valeur du test de fisher",
                                                  "Statistique du test de Kolmogorov",
                                                  "P-valeur du test de Kolmogorov",
                                                  "Statistique du test de Shapiro Wilk",
                                                  "P-valeur du test de Shapiro Wilk",
                                                  "AIC"),
                                    Valeur=c(SCE,score,score_ajuste,Stat_Fisher,p_valeur_F,
                                             stat_kol,p_valeur_kol,stat_sh,p_valeur_sh,AIC)))
  row.names(tableau_stat)=1:10
  tableau_stat
}
###################################################################################################################################
##############fonction estimation droite###########################################################################################
##l'on va formuler la droite de régression et finaliser la liste de tous les résultats
estimationDroite<-function(vary,varx,er=0.05){
  Coefficients=tableauCoef(vary,varx,er) ###dataframe des coefficients
  Tableau_stat=statval(vary,varx,er=0.05)   ####dataframe des statistiques de validation 
  Formule=paste("La droite de regression estimée est: Y =(",Coefficients[,2][1],
                ")+(",Coefficients[,2][2],")X")               #########l'equation du modèle
  resultat=list(Formule,Coefficients,Tableau_stat)
  resultat
}
##################################################################################################################################
#############ajouter un titre######################################################################################################
###fonction de mise en forme des titres dans la sortie excel
xlsx_addTitle<-function(sheet,rowIndex,title,titleStyle){
  rows=createRow(sheet,rowIndex=rowIndex)
  sheetTitle=createCell(row = rows,colIndex = 1)
  setCellValue(sheetTitle[[1,1]],title)
  setCellStyle(sheetTitle[[1,1]],titleStyle)
}
##################################################################################################################################
##################Ecriture de la fontion définitive###############################################################################
#la fonction finale pour faire l'analyse
analyseR<-function(pathfichier,er=0.05,Sortie){
  donnees=Import(pathfichier)
  variables=detectvar(donnees)
  regression=estimationDroite(variables[[1]],variables[[2]],er=0.05)
  statistique=statdesc(donnees)
  ##une fois toutes les statistiques et graphiques construis nous exportons les résutats
  
  if(Sortie=="EXCEL"){
    ##l'exportation nécessitera le package xlsx
    excel=winDialog(type = "yesno","L'exportation de votre fichier excel va nécessiter le package xlsx. L'avez vous déja installé?")
    if (excel=="YES"){
      require(xlsx,quietly = T)
      ##nous créons un classeur exel qui contiendra les résultats
      wb=createWorkbook(type="xlsx")
      ###la mise en forme des cellules, des titres et des sous titres sont ensuite faire
      TitleStyle=CellStyle(wb)+Font(wb,heightInPoints = 16,color="blue",isBold = T,underline = 1)
      sub_titleStyle=CellStyle(wb)+Font(wb,heightInPoints = 14,isItalic = T,isBold = F)
      Table_rowStyle=CellStyle(wb)+Font(wb,isBold = T)
      Table_colStyle=CellStyle(wb)+Font(wb,isBold = T)+
        Alignment(wrapText = T,horizontal = "ALIGN_CENTER")+
        Border(color="black",position=c("TOP","BOTTOM"),pen=c("BORDER_THIN","BORDER_THICK"))
      
      ##la premiere feuille du classeur va contenir le jeu de données qui a été utilisé
      sheet=createSheet(wb,sheetName = "BD_ANALYSE")
      xlsx_addTitle(sheet,rowIndex = 1,title ="Base de données pour l'analyse",titleStyle =TitleStyle)
      addDataFrame(donnees,sheet,startRow = 3,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      setColumnWidth(sheet,colIndex = c(1:ncol(donnees)),colWidth = 11)
      ##puis deux autres feuilles sont crées pour réprésenter les deux histogrammes des variables numériques
      for (i in 1:2) {
        sheet1=createSheet(wb,sheetName =paste("Histogramme",i))
        xlsx_addTitle(sheet1,rowIndex = 1,title =paste("Histogramme",variables[[4]][i]),titleStyle =TitleStyle)
        addPicture(paste("Histogramme",variables[[4]][i],".png"),sheet1,scale=1,
                   startRow = 4,startColumn = 3)
        res=file.remove(paste("Histogramme",variables[[4]][i],".png"))           #parallelement les graphiques qui avaient été sauvegardés
        ##dans le repertoire courant sont supprimés à chaque fois
      }
      ##ensuite deux autres feuilles pour leur boxplot
      for (i in 1:2) {
        sheet2=createSheet(wb,sheetName =paste("Boxplot",i))
        xlsx_addTitle(sheet2,rowIndex = 1,title =paste("Boxplot",variables[[4]][i]),titleStyle =TitleStyle)
        addPicture(paste("Boxplot",variables[[4]][i],".png"),sheet2,scale=1,
                   startRow = 4,startColumn = 3)
        res=file.remove(paste("Boxplot",variables[[4]][i],".png"))
      }
      ###les boxplot multi classes sont sauvegardés dans une feuille chacun
      for (i in 1:2) {
        sheet3=createSheet(wb,sheetName =paste("Boxplot_Multi",i))
        xlsx_addTitle(sheet3,rowIndex = 1,title =paste("Boxplot_multi_cellules",variables[[4]][i]),titleStyle =TitleStyle)
        addPicture(paste("Boxplot_multi_cellules",variables[[4]][i],".png"),sheet3,scale=1,
                   startRow = 4,startColumn = 3)
        res=file.remove(paste("Boxplot_multi_cellules",variables[[4]][i],".png"))
      }
      ################le nuage de points entre les deux variables est également réalisé dans une feuille
      sheet4=createSheet(wb,sheetName = "Nuage_de_points")
      xlsx_addTitle(sheet4,rowIndex = 1,title =paste("Nuage de points de la variable",
                                                     variables[[4]][1],"en fonction de",variables[[4]][2]),
                    titleStyle =TitleStyle) 
      addPicture(paste("Nuage de points de la variable",variables[[4]][1],
                       "en fonction de",variables[[4]][2],".png"),
                 sheet4,scale=1,startRow = 4,startColumn = 3)
      res=file.remove(paste("Nuage de points de la variable",variables[[4]][1],
                            "en fonction de",variables[[4]][2],".png"))
      
      ###une feuille est utilisée pour stocker les statistiques pour la validation du modèle
      sheet6=createSheet(wb,sheetName = "Statistique")
      xlsx_addTitle(sheet6,rowIndex = 1,title ="Résumés descriptifs de la base",titleStyle =TitleStyle)
      xlsx_addTitle(sheet6,rowIndex = 3,title ="Caractéristiques des variables",titleStyle =sub_titleStyle)
      addDataFrame(statistique[[1]],sheet6,startRow =5 ,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      xlsx_addTitle(sheet6,rowIndex = 11,title ="Tableau des effectifs",titleStyle =sub_titleStyle)
      addDataFrame(statistique[[2]],sheet6,startRow =13 ,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      xlsx_addTitle(sheet6,rowIndex = 15+nrow(statistique[[2]]),title ="Coefficients de corrélation linéaire",titleStyle =sub_titleStyle)
      addDataFrame(statistique[[3]],sheet6,startRow =17+nrow(statistique[[2]]) ,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      
      ###########une feuille est utilisée pour présenter les résultats de l'estimation du modèle
      sheet5=createSheet(wb,sheetName = "REGRESSION")
      xlsx_addTitle(sheet5,rowIndex = 1,title ="Résultats de la régression",titleStyle =TitleStyle)
      xlsx_addTitle(sheet5,rowIndex = 3,title ="Coefficients de la régression",titleStyle =sub_titleStyle)
      addDataFrame(regression[[2]],sheet5,startRow = 5,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      xlsx_addTitle(sheet5,rowIndex = 11,title ="Statistiques de validation de la régression",titleStyle =sub_titleStyle)
      addDataFrame(regression[[3]],sheet5,startRow = 13,startColumn = 2,colnamesStyle = Table_colStyle,
                   rownamesStyle = Table_rowStyle)
      xlsx_addTitle(sheet5,rowIndex = 26,title =regression[[1]],titleStyle =sub_titleStyle)
      
      ###########le classeur est ensuite sauvegardé dans un fichier excel
      saveWorkbook(wb,"Resultat_analyseR.xlsx")
    }
  }
  ##########pour l'exportation au format pdf le package grid et gridExtra sont utilisé
  if(Sortie=="PDF"){
    vers_pdf=winDialog(type = "yesno","L'exportation de votre fichier pdf va nécessiter le package gridExtra. L'avez vous déja installé?")
    if (vers_pdf=="YES"){
      require(gridExtra,quietly = T)
      require(grid,quietly = T)
      pdf(file="Resultat_analyseR.pdf",paper = "a4") ###le pdf est crée
      grid.table(head(donnees))   ##on sauvegarde les premières lignes du jeu de données sur une page
      grid.newpage()
      
      plot(variables[[1]]~variables[[2]],col=as.numeric(factor(variables[[3]])),
           xlab = variables[[4]][2],ylab = variables[[4]][1],
           main=paste("Nuage de points de la variable",variables[[4]][1],
                      "en fonction de",variables[[4]][2]))                   ##puis le nuage de point est représenté sur la page suivante
      grid.newpage()
      
      par(mfrow=c(2,1)) ###on représente ensuite les boxplots de chaque variable numériques
      par(mfg=c(1,1))
      boxplot(variables[[1]],col = "Blue",xlab =variables[[4]][1],
              main=paste("Boxplot",variables[[4]][1]))
      par(mfg=c(2,1))
      boxplot(variables[[2]],col = "Blue",xlab =variables[[4]][2],
              main=paste("Boxplot",variables[[4]][2]))
      grid.newpage()
      
      par(mfrow=c(2,1))#########on représente les boxplot multicellules
      par(mfg=c(1,1))
      boxplot(variables[[1]]~factor(variables[[3]]),col=1:24,xlab = variables[[4]][3],ylab = variables[[4]][1],
              main=paste("Boxplot_multi_cellules",variables[[4]][1]))
      par(mfg=c(2,1))
      boxplot(variables[[2]]~factor(variables[[3]]),col=1:24,xlab = variables[[4]][3],ylab = variables[[4]][2],
              main=paste("Boxplot_multi_cellules",variables[[4]][2]))
      grid.newpage()
      #########on ecrit les résultes de l'estimation du modèle sur une autre page 
      grid.table(regression[[1]])
      grid.newpage()
      
      grid.table(regression[[2]])
      grid.newpage()
      
      grid.table(regression[[3]])
      grid.newpage()
      
      grid.table(statistique[[1]])
      grid.newpage()
      
      grid.table(statistique[[3]])
      grid.newpage()
      
      grid.table(statistique[[2]])
      grid.newpage()
      
      par(mfrow=c(2,1)) #########################on ecrit enfin les histogrammes
      par(mfg=c(1,1))
      hist(variables[[1]],col = "yellow", probability = T,main = paste("Histogram of",variables[[4]][1]),xlab =variables[[4]][1])
      lines(density(variables[[1]]),col="blue",lty="solid",lwd=1)
      par(mfg=c(2,1))
      hist(variables[[2]],col = "yellow", probability = T,main = paste("Histogram of",variables[[4]][2]),xlab =variables[[4]][2])
      lines(density(variables[[2]]),col="blue",lty="solid",lwd=1)
      ###on ferme la fenetre graphique permettant de remplir le pdf
      suppressMessages(dev.off())
    }
  }
##Ainsi s'achève notre programme
}