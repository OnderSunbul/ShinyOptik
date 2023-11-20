
okuryazar<-function(){
  
  defaultW <- getOption("warn") 
  
  options(warn = -1) 


### PAKETLER####

require(stringr)
require(openxlsx)
require(psych)
require(itemanalysis)
  require(crayon)
  require(ggplot2)
#### GENEL BİLGİLER ### Bakılacak 



##### TAM PUAN GİRİŞİ ######
tam_puan <- readline(prompt = "Tam puan giriniz: ")
tam_puan <- as.integer(tam_puan)



seceneksayisi <-5

options<-c("A","B","C","D","E")


#tam_puan <-100



#### MADDE ANALİZİ FONKSİYONU#####



ITEMAN3 <-
  function (data=secenek,
            key=anahtar,
            options=options,
            ngroup = ncol(data) + 1,
            correction = TRUE)
  {
    #########################################################################################
    # data, a data frame with N rows and n columns, where N denotes the number of subjects
    #        and n denotes the number of items. All items should be scored using nominal
    #        response categories. All variables (columns) must be "character".
    #        Missing values ("NA") are allowed and scored as incorrect in item analysis
    
    # key,     a character vector of length n, where n denotes the number of items.
    
    # options, number of possible nominal options for items (e.g., "A","B","C","D")
    #          make sure each item is consistent, and includes the same response options
    
    # ngroup, number of score groups
    #
    # correction, TRUE or FALSE, if TRUE item and distractor discrimination is corrected for
    # 		  spuriousnes by removing the item score from the total score
    #########################################################################################
    
    
    for (i in 1:ncol(data)) {
      if (is.character(data[, i]) != TRUE) {
        data[, i] = as.character(data[, i])
      }
    }
    
    scored.data <-
      as.data.frame(matrix(nrow = nrow(data), ncol = ncol(data)))
    
    for (i in 1:ncol(scored.data)) {
      scored.data[, i] <- ifelse(data[, i] == key[i], 1, 0)
      if (length(which(is.na(scored.data[, i]))) != 0) {
        scored.data[which(is.na(scored.data[, i]) == TRUE), i] = 0
      }
    }
    
    total.score <- rowSums(scored.data)
    ybar <- mean(total.score)
    sdt <- sd(total.score)
    p <- colMeans(scored.data)
    
    pbis <- c()
    pbis.corrected <- c()
    bis  <- c()
    bis.corrected <- c()
    
    for (k in 1:ncol(data)) {
      pbis[k] = cor(scored.data[, k], total.score, use = "pairwise.complete.obs")
      pbis.corrected[k] = cor(scored.data[, k],
                              rowMeans(scored.data[, -k], na.rm = TRUE) * (ncol(scored.data) -
                                                                             1),
                              use = "pairwise.complete.obs")
      bis[k] = polyserial(total.score, scored.data[, k])
      bis.corrected[k] = polyserial(rowMeans(scored.data[, -k], na.rm = TRUE) *
                                      (ncol(scored.data) - 1),
                                    scored.data[, k])
    }
    
    
    item.stat <- matrix(nrow = ncol(data), ncol = 4)
    colnames(item.stat) <-
      c("Madde Güçlüğü",
        "Madde Eşik Değeri",
        "Point-Biserial",
        "Ayırt Edicilik")
    
    # rnames <- ("Madde 1")
    # for (i in 2:ncol(data)) {
    #   rnames <- c(rnames, paste("Madde ", i, sep = ""))
    # }
    
    rnames <-colnames(secenek)
    
    rownames(item.stat) <- rnames
    
    
    
    item.stat[, 1] = p
    item.stat[, 2] = qnorm(1 - p)
    if (correction == TRUE) {
      item.stat[, 3] = pbis.corrected
    } else {
      item.stat[, 3] = pbis
    }
    if (correction == TRUE) {
      item.stat[, 4] = bis.corrected
    } else {
      item.stat[, 4] = bis
    }
    
    
    return(round(item.stat[,c(1,4)],2))
    
    
  } 

veriyolu <- file.choose()
okuma120 <-
  read.fwf(
    veriyolu,
    header = F,
    stringsAsFactors = F,
    widths = c(11, 12, 12,11,3,1,1,1,1,1,1,1,120),
    fileEncoding = "latin5"
  )


tam_okuma<-okuma120



madde_sayisi <-
  str_count(sub('    *\\    ', '', okuma120[1, 13]))# cevap anahtarında boşluk olan durum için kontrol edilmeli



tam_okuma <-
  read.fwf(
    veriyolu,
    header = F,
    widths = c(11, 12, 12,11,3,1,1,1,1,1,1,1, rep(1, madde_sayisi)),
    fileEncoding = "latin5",
    stringsAsFactors = F,
  )


names(tam_okuma)[1] ="BinNo"
names(tam_okuma)[2] = "Ad"
names(tam_okuma)[3] = "Soyad"
names(tam_okuma)[4] = "ÖgrenciNumarası"
names(tam_okuma)[5] = "GrupNumarası"
names(tam_okuma)[6] = "ÖğretimTürü"
names(tam_okuma)[7] = "Sınıf"
names(tam_okuma)[8] = "SınavTürü"
names(tam_okuma)[9] = "SınavDönemi"
names(tam_okuma)[10] = "Sınavİptali"
names(tam_okuma)[11] = "SınavaKatılmama"
names(tam_okuma)[12] = "KitapçıkTürü"

tam_okuma$ÖgrenciNumarası<-as.numeric(tam_okuma$ÖgrenciNumarası)

names(tam_okuma)[13:(dim(tam_okuma)[2])] <-
  paste0("M", 1:madde_sayisi)





tam_okumaKatılmayanlar<-tam_okuma[which(tam_okuma$SınavaKatılmama =="K"),]
tam_okumaİptalEdilenler<-tam_okuma[which(tam_okuma$Sınavİptali=="İ"),]

#tamveri<-tam_okuma[which(tam_okuma$SınavaKatılmama !="K"),]# Katılmayanlar Çıkartılıyor


tamveri<-tam_okuma


#tamveri<-tamveri[which(tamveri$Sınavİptali!="İ"),] # İptal Edilenler Çıkartılıyor

#KatıldığıHaldeKatılmadıİşaretlenenler<-tamveri[which(tamveri$SınavaKatılmama =="K"),]
KHKİ<-tam_okuma[which(tam_okuma$SınavaKatılmama =="K"),]


#KHKİ[,13:62][sapply(KHKİ[,13:62], function(x) as.character(x)==" ")] <- NA


KHKİ<-KHKİ[,13:(12+madde_sayisi)][rowSums(is.na(KHKİ[,13:(12+madde_sayisi)])) < 1, ]


KATHOK<-tam_okuma[which(tam_okuma$SınavaKatılmama !="K"),]


#KATHOK[,13:62][sapply(KATHOK[,13:62], function(x) as.character(x)==" ")] <- NA


KATHOK<-KATHOK[,13:(12+madde_sayisi)][rowSums(is.na(KATHOK[,13:(12+madde_sayisi)])) >= madde_sayisi, ]












if (tam_okuma[1, 12] == "A" &
    tam_okuma[2, 12] == "B"&
    !is.na(tam_okuma[1, 12]))
{

  
  ##### KİTAPÇIK AYRIŞIMI ######
  
  #tam_okuma$"Kitapçık Türü"<-as.character(tam_okuma$"Kitapçık Türü")
  
  Kitapçıkİşaretlemeyenler<-tamveri[which(tamveri$KitapçıkTürü ==" "||tamveri$KitapçıkTürü ==""||tamveri$KitapçıkTürü =="NA"),]
  
  if(nrow(Kitapçıkİşaretlemeyenler)==0){
    
   cat(green("SINAVDA BÜTÜN ADAYLAR KİTAPÇIK TÜRÜNÜ İŞARETLEMİŞTİR\n", "\n")) 
    
  }
  
  if(nrow(Kitapçıkİşaretlemeyenler)!=0){
    
    cat(red("SINAVDA  KİTAPÇIK TÜRÜNÜ İŞARETLEMEYEN ADAYLAR BULUNMAKTADIR"), "\n")
    print(Kitapçıkİşaretlemeyenler[1:5])
    
  }
  
  
  
  
  tam_okumaA<-tamveri[which(tamveri$KitapçıkTürü =="A"),]
  
  tam_okumaB<-tamveri[which(tamveri$KitapçıkTürü =="B"),]



#### İPTAL KARARI SEÇENEKLİ  #####



iptal_edilecek_soru_kararı <-
  menu(c("EVET", "HAYIR"), title = "İPTAL EDİLECEK SORU VAR MI?")



if (iptal_edilecek_soru_kararı == 2) {



####  VERİ OKUMA  #### 



##BİLGİ BASIMI #####

cat("SINAVDA ", red(paste0(dim(tam_okuma)[2] - 12)), " SORU YER ALMAKTADIR", "\n")

cat("SINAVDA ", red("A - B")," OLMAK ÜZERE 2 KİTAPÇIK TÜRÜ BULUNMAKTADIR", "\n")
  
cat("SINAVDA ", red(paste0(nrow(tam_okuma)-2)), " OPTİK KAĞIT İŞLEM GÖRMÜŞTÜR", "\n")

cat("SINAVA ", red(paste0(nrow(tam_okumaKatılmayanlar))), " ADAYIN KATILMADIĞI RAPORLANMIŞTIR", "\n")


cat("SINAVDA ", red(paste0(nrow(tam_okumaİptalEdilenler))), " ADAYIN SINAVI İPTAL EDİLMİŞTİR", "\n")


cat("SINAVDA ", red(paste0(nrow(KHKİ))), " ADAYIN FORMU ADAY SINAVA KATILDIĞI HALDE KATILMADI İŞARETLENMİŞTİR", "\n")


cat("SINAVDA ", red(paste0(nrow(KATHOK))), " ADAYIN FORMU ADAY SINAVA KATILMADIĞI HALDE KATILDI İŞARETLENMİŞTİR", "\n")


##### PUANLAMA BİLGİLERİ ###### BAKILACAK


puanlamabilgileriA <- tamveri[which(tamveri$KitapçıkTürü =="A"),]
puanlamabilgileriB <- tamveri[which(tamveri$KitapçıkTürü =="B"),]


puanlamabilgileri<-rbind(puanlamabilgileriA,puanlamabilgileriB)

#####



dy_matrisiA <-
  matrix(NA,nrow= (dim(tam_okumaA)[1]), ncol=(dim(tam_okumaA)[2] )-12)



dy_matrisiB<-
  matrix(NA,nrow= (dim(tam_okumaB)[1]), ncol=(dim(tam_okumaB)[2] )-12)


# A Kitapçığı 
for (i in 1:dim(tam_okumaA)[1]) {
  
  dy_matrisiA[i, ] <-
    ifelse(tam_okumaA[i, 13:dim(tam_okumaA)[2]] == tam_okumaA[1, 13:dim(tam_okumaA)[2]], 1, 0)
  #Soru İptalleri için düzeltme
  
  
  
}
#Soru İptalleri için düzeltme
#dy_matrisia[,c(4,7,30,34)]<-1


# B Kitapçığı
for (i in 1:dim(tam_okumaB)[1]) {
  
  dy_matrisiB[i, ] <-
    ifelse(tam_okumaB[i, 13:dim(tam_okumaB)[2]] == tam_okumaB[1, 13:dim(tam_okumaB)[2]], 1, 0)
  #Soru İptalleri için düzeltme
  
  
  
}
#Soru İptalleri için düzeltme
#dy_matrisib[,c(45,46,21,17)]<-1

dy_matrisi<-rbind(dy_matrisiA,dy_matrisiB)





anahtara<- as.character(as.vector(tam_okumaA[1,-c(1:12)]))   

anahtarb<- as.character(as.vector(tam_okumaB[2,-c(1:12)]))   


#seceneka<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="A"),]

secenekA<-tam_okumaA[-1,-c(1:12)]


#secenekb<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="B"),]

secenekB<-tam_okumaB[-1,-c(1:12)]






dy10 <- apply(dy_matrisi, 2, as.numeric)

dsay <- apply(dy10, 1, sum)



Puan <- round(dsay * tam_puan /madde_sayisi, 0)






### PUAN BETİMLEME İÇİN PUAN


dy_matrisiP<-rbind(dy_matrisiA[-1,],dy_matrisiB[-1,])





dy10P <- apply(dy_matrisiP, 2, as.numeric)

dsayP <- apply(dy10P, 1, sum)



PuanP <- round(dsayP * tam_puan / madde_sayisi, 0)
############################################################


yazılacakdosya0 <-
  data.frame(cbind(puanlamabilgileri[, 1:12], dsay, Puan))
names(yazılacakdosya0)[13] <- "Doğru Cevap Sayısı"
yazılacakdosya1 <-
  yazılacakdosya0[order(yazılacakdosya0$Puan,decreasing = T), ]


yazılacakdosya1$SınavaKatılmama<-"KATILDI"

### KATILMAYAN DÜZENLEMELERİ   ###
# tam_okumaKatılmayanlar$SınavaKatılmama<-"KATILMADI"
# tam_okumaKatılmayanlar1<-tam_okumaKatılmayanlar[c(1:12)]
# tam_okumaKatılmayanlar1$'Doğru Cevap Sayısı'<- "-"
# tam_okumaKatılmayanlar1$Puan<- "-"

### İPTAL EDİLEN DÜZENLEMELERİ ######

# 
# tam_okumaİptalEdilenler$Sınavİptali<-"İPTAL EDİLDİ"
# tam_okumaİptalEdilenler1<-tam_okumaİptalEdilenler[c(1:12)]
# tam_okumaİptalEdilenler1$'Doğru Cevap Sayısı'<- "-"
# tam_okumaİptalEdilenler1$Puan<- "-"





#yazdırılacak2<-rbind(yazılacakdosya1,tam_okumaKatılmayanlar1,tam_okumaİptalEdilenler1)
  
yazdırılacak2<-yazılacakdosya1

####### MADDE ANALİZİ ##############





secenek<-secenekA

maddeanalizia<-ITEMAN3(key = anahtara)


secenek<-secenekB

maddeanalizib<-ITEMAN3(key = anahtarb)



### YAZDIRMA İŞLEMLERİ #####




hs <-
  createStyle(
    textDecoration = "BOLD",
    fontColour = "#FFFFFF",
    fontSize = 12,
    fontName = "Arial Narrow",
    fgFill = "#4F80BD"
  )

x <- chartr("\\", "/", veriyolu)


dosyaadı <- tail(strsplit(x, "/")[[1]], 1)



wb <- createWorkbook()
addWorksheet(wb, "SONUÇLAR")
addWorksheet(wb, "GRAFİK", gridLines=FALSE)
addWorksheet(wb, "BETİMSEL", gridLines=T)
addWorksheet(wb, "MADDE ANALİZİA", gridLines=T)
addWorksheet(wb, "MADDE ANALİZİB", gridLines=T)





writeData(wb, "SONUÇLAR", yazdırılacak2,
          colNames = TRUE,
          borders = "all",
          #firstActiveRow = 2,
          headerStyle = hs,
          withFilter = T
          #overwrite = F
)


setColWidths(wb, sheet = "SONUÇLAR", cols = 1:13, widths = "auto")

freezePane(wb, "SONUÇLAR", firstRow = T)

betimsel<-round(describe(PuanP),2)

betimsel1<-betimsel[,c(2,3,4,5,8,9)]

colnames(betimsel1)<-c("Birey_Sayısı","Ortalama_Puan","S.Sapma","Ortanca","Min","Max")





writeData(wb, "BETİMSEL",betimsel1 ,
          colNames = TRUE,
          borders = "surrounding",
          #firstActiveRow = 2,
          headerStyle = hs,
          #overwrite = F
)

setColWidths(wb, sheet = "BETİMSEL", cols = 1:6, widths = "auto")







png("graph.png", width=1024, height=768, units="px", res=144)  #output to png device
p <- hist(Puan,main="ÖLÇME ve DEĞERLENDİRME MERKEZİ",col = "red",ylab = "FREKANS", xlim = c(0,100))
abline(v=60,col="green",lwd=3, lty=2)
#print(p)
dev.off()  #important to shut down the active png device
insertImage(wb, "GRAFİK", "graph.png", width=11.18, height=7.82, units="in")




writeData(wb, "MADDE ANALİZİA", maddeanalizia ,
          colNames = TRUE,
          rowNames = T,
          borders = "all",
          #firstActiveRow = 2,
          headerStyle = hs,
          #overwrite = F
)

setColWidths(wb, sheet = "MADDE ANALİZİA", cols = 1:3, widths = c(15,15,15))




writeData(wb, "MADDE ANALİZİB", maddeanalizib ,
          colNames = TRUE,
          rowNames = T,
          borders = "all",
          #firstActiveRow = 2,
          headerStyle = hs,
          #overwrite = F
)

setColWidths(wb, sheet = "MADDE ANALİZİB", cols = 1:3, widths = c(15,15,15))





saveWorkbook(wb, file=choose.files(caption="KAYDETMEK İSTEDİĞİNİZ DOSYA ADINI YAZINIZ", 
                                   filters = c("Excel Dosyaları (.xlsx)","*.xlsx")), overwrite=TRUE)


}



if (iptal_edilecek_soru_kararı == 1) {
  soru_iptalleriA <-
    readline(prompt = "A KİTAPÇIĞINDA İPTALİ İSTENEN SORULARI ARALARINA VİRGÜL KOYARAK GİRİNİZ: ")
  iptallerA <- as.integer(unlist(strsplit(soru_iptalleriA, ",")))
  
  soru_iptalleriB <-
    readline(prompt = "B KİTAPÇIĞINDA İPTALİ İSTENEN SORULARI ARALARINA VİRGÜL KOYARAK GİRİNİZ: ")
  iptallerB <- as.integer(unlist(strsplit(soru_iptalleriB, ",")))
  
  
  
  
  tam_okumaA <- tam_okumaA[, -(12 + c(iptallerA))]
  
  
  tam_okumaB <- tam_okumaB[, -(12 + c(iptallerB))]
  
  
  
  
  cat("İPTALLERDEN SONRA SINAVDA ",
      paste0(dim(tam_okumaA)[2] - 12),
      " SORU KALMIŞTIR",
      "\n")
  
  
  
  ##### PUANLAMA BİLGİLERİ ###### BAKILACAK
  
  
  puanlamabilgileriA <- tamveri[which(tamveri$KitapçıkTürü =="A"),]
  puanlamabilgileriB <- tamveri[which(tamveri$KitapçıkTürü =="B"),]
  
  
  puanlamabilgileri<-rbind(puanlamabilgileriA,puanlamabilgileriB)
  
  #####
  
  
  
  dy_matrisiA <-
    matrix(NA,nrow= (dim(tam_okumaA)[1]), ncol=(dim(tam_okumaA)[2] )-12)
  
  
  
  dy_matrisiB<-
    matrix(NA,nrow= (dim(tam_okumaB)[1]), ncol=(dim(tam_okumaB)[2] )-12)
  
  
  # A Kitapçığı 
  for (i in 1:dim(tam_okumaA)[1]) {
    
    dy_matrisiA[i, ] <-
      ifelse(tam_okumaA[i, 13:dim(tam_okumaA)[2]] == tam_okumaA[1, 13:dim(tam_okumaA)[2]], 1, 0)
    #Soru İptalleri için düzeltme
    
    
    
  }
  #Soru İptalleri için düzeltme
  #dy_matrisia[,c(4,7,30,34)]<-1
  
  
  # B Kitapçığı
  for (i in 1:dim(tam_okumaB)[1]) {
    
    dy_matrisiB[i, ] <-
      ifelse(tam_okumaB[i, 13:dim(tam_okumaB)[2]] == tam_okumaB[1, 13:dim(tam_okumaB)[2]], 1, 0)
    #Soru İptalleri için düzeltme
    
    
    
  }
  #Soru İptalleri için düzeltme
  #dy_matrisib[,c(45,46,21,17)]<-1
  
  dy_matrisi<-rbind(dy_matrisiA,dy_matrisiB)
  
  
  
  
  
  anahtara<- as.character(as.vector(tam_okumaA[1,-c(1:12)]))   
  
  anahtarb<- as.character(as.vector(tam_okumaB[2,-c(1:12)]))   
  
  
  #seceneka<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="A"),]
  
  secenekA<-tam_okumaA[-1,-c(1:12)]
  
  
  #secenekb<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="B"),]
  
  secenekB<-tam_okumaB[-1,-c(1:12)]
  
  
  
  
  
  
  dy10 <- apply(dy_matrisi, 2, as.numeric)
  
  dsay <- apply(dy10, 1, sum)
  
  
  
  Puan <- round(dsay * tam_puan /(madde_sayisi-length(iptallerA)), 0)
  
  
  
  
 
  
  ### PUAN BETİMLEME İÇİN PUAN
  
  
  dy_matrisiP<-rbind(dy_matrisiA[-1,],dy_matrisiB[-1,])
  
  
  
  
  
  dy10P <- apply(dy_matrisiP, 2, as.numeric)
  
  dsayP <- apply(dy10P, 1, sum)
  
  
  
  PuanP <- round(dsayP * tam_puan / madde_sayisi, 0)
  ############################################################
  
  
  yazılacakdosya0 <-
    data.frame(cbind(puanlamabilgileri[, 1:12], dsay, Puan))
  names(yazılacakdosya0)[13] <- "Doğru Cevap Sayısı"
  yazılacakdosya1 <-
    yazılacakdosya0[order(yazılacakdosya0$Puan,decreasing = T), ]
  
  ### KATILMA DURUMU DÜZENLEMELERİ   #####
  
  yazılacakdosya1$SınavaKatılmama<-"KATILDI"
  # tam_okumaKatılmayanlar$SınavaKatılmama<-"KATILMADI"
  # tam_okumaKatılmayanlar1<-tam_okumaKatılmayanlar[c(1:12)]
  # tam_okumaKatılmayanlar1$'Doğru Cevap Sayısı'<- "-"
  # tam_okumaKatılmayanlar1$Puan<- "-"
  # 
  
  ### İPTAL EDİLEN DÜZENLEMELERİ ######
  
  # 
  # tam_okumaİptalEdilenler$Sınavİptali<-"İPTAL EDİLDİ"
  # tam_okumaİptalEdilenler1<-tam_okumaİptalEdilenler[c(1:12)]
  # tam_okumaİptalEdilenler1$'Doğru Cevap Sayısı'<- "-"
  # tam_okumaİptalEdilenler1$Puan<- "-"
  # 
  # 
  # 
  
  
  #yazdırılacak2<-rbind(yazılacakdosya1,tam_okumaKatılmayanlar1,tam_okumaİptalEdilenler1)
  
  yazdırılacak2<-yazılacakdosya1
  
  ####### MADDE ANALİZİ ##############
  
  
  
  
  
  secenek<-secenekA
  
  maddeanalizia<-ITEMAN3(key = anahtara)
  
  
  secenek<-secenekB
  
  maddeanalizib<-ITEMAN3(key = anahtarb)
  
  
  
  ### YAZDIRMA İŞLEMLERİ #####
  
  
  
  
  hs <-
    createStyle(
      textDecoration = "BOLD",
      fontColour = "#FFFFFF",
      fontSize = 12,
      fontName = "Arial Narrow",
      fgFill = "#4F80BD"
    )
  
  x <- chartr("\\", "/", veriyolu)
  
  
  dosyaadı <- tail(strsplit(x, "/")[[1]], 1)
  
  
  
  wb <- createWorkbook()
  addWorksheet(wb, "SONUÇLAR")
  addWorksheet(wb, "GRAFİK", gridLines=FALSE)
  addWorksheet(wb, "BETİMSEL", gridLines=T)
  addWorksheet(wb, "MADDE ANALİZİA", gridLines=T)
  addWorksheet(wb, "MADDE ANALİZİB", gridLines=T)
  
  
  
  
  
  writeData(wb, "SONUÇLAR", yazdırılacak2,
            colNames = TRUE,
            borders = "all",
            #firstActiveRow = 2,
            headerStyle = hs,
            withFilter = T
            #overwrite = F
  )
  
  
  setColWidths(wb, sheet = "SONUÇLAR", cols = 1:13, widths = "auto")
  
  freezePane(wb, "SONUÇLAR", firstRow = T)
  
  betimsel<-round(describe(PuanP),2)
  
  betimsel1<-betimsel[,c(2,3,4,5,8,9)]
  
  colnames(betimsel1)<-c("Birey_Sayısı","Ortalama_Puan","S.Sapma","Ortanca","Min","Max")
  
  
  
  
  
  writeData(wb, "BETİMSEL",betimsel1 ,
            colNames = TRUE,
            borders = "surrounding",
            #firstActiveRow = 2,
            headerStyle = hs,
            #overwrite = F
  )
  
  setColWidths(wb, sheet = "BETİMSEL", cols = 1:6, widths = "auto")
  
  
  
  
  
  
  
  png("graph.png", width=1024, height=768, units="px", res=144)  #output to png device
  p <- hist(Puan,main="ÖLÇME ve DEĞERLENDİRME MERKEZİ",col = "red",ylab = "FREKANS", xlim = c(0,100))
  abline(v=60,col="green",lwd=3, lty=2)
  #print(p)
  dev.off()  #important to shut down the active png device
  insertImage(wb, "GRAFİK", "graph.png", width=11.18, height=7.82, units="in")
  
  
  
  
  writeData(wb, "MADDE ANALİZİA", maddeanalizia ,
            colNames = TRUE,
            rowNames = T,
            borders = "all",
            #firstActiveRow = 2,
            headerStyle = hs,
            #overwrite = F
  )
  
  setColWidths(wb, sheet = "MADDE ANALİZİA", cols = 1:3, widths = c(15,15,15))
  
  
  
  
  writeData(wb, "MADDE ANALİZİB", maddeanalizib ,
            colNames = TRUE,
            rowNames = T,
            borders = "all",
            #firstActiveRow = 2,
            headerStyle = hs,
            #overwrite = F
  )
  
  setColWidths(wb, sheet = "MADDE ANALİZİB", cols = 1:3, widths = c(15,15,15))
  
  
  
  
  
  saveWorkbook(wb, file=choose.files(caption="KAYDETMEK İSTEDİĞİNİZ DOSYA ADINI YAZINIZ", 
                                     filters = c("Excel Dosyaları (.xlsx)","*.xlsx")), overwrite=TRUE)
  
  
  
  
  
  
}


} else{
  
  




    #tam_okumaA<-tamveri[which(tamveri$KitapçıkTürü =="A"),]
    
    #tam_okumaB<-tamveri[which(tamveri$KitapçıkTürü =="B"),]
    ##BİLGİ BASIMI #####
  
  ##BİLGİ BASIMI #####
  
  cat("SINAVDA ", red(paste0(dim(tam_okuma)[2] - 12)), " SORU YER ALMAKTADIR", "\n")
  
  cat("SINAVDA ", red("1"),"KİTAPÇIK TÜRÜ BULUNMAKTADIR", "\n")
  
  cat("SINAVDA ", red(paste0(nrow(tam_okuma)-1)), " OPTİK KAĞIT İŞLEM GÖRMÜŞTÜR", "\n")
  
  cat("SINAVA ", red(paste0(nrow(tam_okumaKatılmayanlar))), " ADAYIN KATILMADIĞI RAPORLANMIŞTIR", "\n")
  
  
  cat("SINAVDA ", red(paste0(nrow(tam_okumaİptalEdilenler))), " ADAYIN SINAVI İPTAL EDİLMİŞTİR", "\n")
  
  
  cat("SINAVDA ", red(paste0(nrow(KHKİ))), " ADAYIN FORMU ADAY SINAVA KATILDIĞI HALDE KATILMADI İŞARETLENMİŞTİR", "\n")
  
  
  cat("SINAVDA ", red(paste0(nrow(KATHOK))), " ADAYIN FORMU ADAY SINAVA KATILMADIĞI HALDE KATILDI İŞARETLENMİŞTİR", "\n")
  
  
    ##### PUANLAMA BİLGİLERİ ###### BAKILACAK
    
    
    #puanlamabilgileriA <- tamveri[which(tamveri$KitapçıkTürü =="A"),]
    #puanlamabilgileriB <- tamveri[which(tamveri$KitapçıkTürü =="B"),]
    
    
    #puanlamabilgileri<-rbind(puanlamabilgileriA,puanlamabilgileriB)
    
    puanlamabilgileri<-tamveri
    


  
  
  #### İPTAL KARARI SEÇENEKLİ  #####
  
  
  
  iptal_edilecek_soru_kararı <-
    menu(c("EVET", "HAYIR"), title = "İPTAL EDİLECEK SORU VAR MI?")
  
  
  
  if (iptal_edilecek_soru_kararı == 2) {
    
    
    
    ####  VERİ OKUMA  #### 
    
    
    
    ##### KİTAPÇIK AYRIŞIMI ######
    
    
    #####
    
    
    
    #dy_matrisiA <-
     # matrix(NA,nrow= (dim(tam_okumaA)[1]), ncol=(dim(tam_okumaA)[2] )-12)
    
    
    
    #dy_matrisiB<-
      #matrix(NA,nrow= (dim(tam_okumaB)[1]), ncol=(dim(tam_okumaB)[2] )-12)
    
    
    dy_matrisi <-
     matrix(NA,nrow= (dim(tamveri)[1]), ncol=(dim(tamveri)[2] )-12)
    
    
    
    
    
    #  Kitapçığı 
    for (i in 1:dim(tamveri)[1]) {
      
      dy_matrisi[i, ] <-
        ifelse(tamveri[i, 13:dim(tamveri)[2]] == tamveri[1, 13:dim(tamveri)[2]], 1, 0)
      #Soru İptalleri için düzeltme
      
      
      
    }
    #Soru İptalleri için düzeltme
    #dy_matrisia[,c(4,7,30,34)]<-1
    
    
    # # B Kitapçığı
    # for (i in 1:dim(tam_okumaB)[1]) {
    #   
    #   dy_matrisiB[i, ] <-
    #     ifelse(tam_okumaB[i, 13:dim(tam_okumaB)[2]] == tam_okumaB[1, 13:dim(tam_okumaB)[2]], 1, 0)
    #   #Soru İptalleri için düzeltme
    #   
    #   
    #   
    # }
    #Soru İptalleri için düzeltme
    #dy_matrisib[,c(45,46,21,17)]<-1
    
   # dy_matrisi<-rbind(dy_matrisiA,dy_matrisiB)
    
    
    
    
    
    anahtar<- as.character(as.vector(tamveri[1,-c(1:12)]))   
    
    #anahtarb<- as.character(as.vector(tam_okumaB[2,-c(1:12)]))   
    
    
    #seceneka<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="A"),]
    
    secenek<-tamveri[-1,-c(1:12)]
    
    
    #secenekb<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="B"),]
    
    #secenekB<-tam_okumaB[-1,-c(1:12)]
    
    
    
    
    
    
    dy10 <- apply(dy_matrisi, 2, as.numeric)
    
    dsay <- apply(dy10, 1, sum)
    
    
    
    Puan <- round(dsay * tam_puan /madde_sayisi, 0)
    
    
    
    
    
    
    ### PUAN BETİMLEME İÇİN PUAN
    
    
    dy_matrisiP<-dy_matrisi[-1,]
    
    
    
    
    
    dy10P <- apply(dy_matrisiP, 2, as.numeric)
    
    dsayP <- apply(dy10P, 1, sum)
    
    
    
    PuanP <- round(dsayP * tam_puan / madde_sayisi, 0)
    ############################################################
    
    
    yazılacakdosya0 <-
      data.frame(cbind(puanlamabilgileri[, 1:12], dsay, Puan))
    names(yazılacakdosya0)[13] <- "Doğru Cevap Sayısı"
    yazılacakdosya1 <-
      yazılacakdosya0[order(yazılacakdosya0$Puan,decreasing = T), ]
    
    
    yazılacakdosya1$SınavaKatılmama<-"KATILDI"
    
    ### KATILMAYAN DÜZENLEMELERİ   ###
    # tam_okumaKatılmayanlar$SınavaKatılmama<-"KATILMADI"
    # tam_okumaKatılmayanlar1<-tam_okumaKatılmayanlar[c(1:12)]
    # tam_okumaKatılmayanlar1$'Doğru Cevap Sayısı'<- "-"
    # tam_okumaKatılmayanlar1$Puan<- "-"
    # 
    ### İPTAL EDİLEN DÜZENLEMELERİ ######
    
    # 
    # tam_okumaİptalEdilenler$Sınavİptali<-"İPTAL EDİLDİ"
    # tam_okumaİptalEdilenler1<-tam_okumaİptalEdilenler[c(1:12)]
    # tam_okumaİptalEdilenler1$'Doğru Cevap Sayısı'<- "-"
    # tam_okumaİptalEdilenler1$Puan<- "-"
    # 
    
    
    
    
   # yazdırılacak2<-rbind(yazılacakdosya1,tam_okumaKatılmayanlar1,tam_okumaİptalEdilenler1)
    
    yazdırılacak2<-yazılacakdosya1
    
    ####### MADDE ANALİZİ ##############
    
    
    
    
    
    #secenek<-secenekA
    
    maddeanalizi<-ITEMAN3(key = anahtar)
    
    # 
    # secenek<-secenekB
    # 
    # maddeanalizib<-ITEMAN3(key = anahtarb)
    # 
    
    
    ### YAZDIRMA İŞLEMLERİ #####
    
    
    
    
    hs <-
      createStyle(
        textDecoration = "BOLD",
        fontColour = "#FFFFFF",
        fontSize = 12,
        fontName = "Arial Narrow",
        fgFill = "#4F80BD"
      )
    
    x <- chartr("\\", "/", veriyolu)
    
    
    dosyaadı <- tail(strsplit(x, "/")[[1]], 1)
    
    
    
    wb <- createWorkbook()
    addWorksheet(wb, "SONUÇLAR")
    addWorksheet(wb, "GRAFİK", gridLines=FALSE)
    addWorksheet(wb, "BETİMSEL", gridLines=T)
    addWorksheet(wb, "MADDE ANALİZİ", gridLines=T)
    #addWorksheet(wb, "MADDE ANALİZİB", gridLines=T)
    
    
    
    
    
    writeData(wb, "SONUÇLAR", yazdırılacak2,
              colNames = TRUE,
              borders = "all",
              #firstActiveRow = 2,
              headerStyle = hs,
              withFilter = T
              #overwrite = F
    )
    
    
    setColWidths(wb, sheet = "SONUÇLAR", cols = 1:13, widths = "auto")
    
    freezePane(wb, "SONUÇLAR", firstRow = T)
    
    betimsel<-round(describe(PuanP),2)
    
    betimsel1<-betimsel[,c(2,3,4,5,8,9)]
    
    colnames(betimsel1)<-c("Birey_Sayısı","Ortalama_Puan","S.Sapma","Ortanca","Min","Max")
    
    
    
    
    
    writeData(wb, "BETİMSEL",betimsel1 ,
              colNames = TRUE,
              borders = "surrounding",
              #firstActiveRow = 2,
              headerStyle = hs,
              #overwrite = F
    )
    
    setColWidths(wb, sheet = "BETİMSEL", cols = 1:6, widths = "auto")
    
    
    
    
    
    
    
    png("graph.png", width=1024, height=768, units="px", res=144)  #output to png device
    p <- hist(Puan,main="ÖLÇME ve DEĞERLENDİRME MERKEZİ",col = "red",ylab = "FREKANS", xlim = c(0,100))
    abline(v=60,col="green",lwd=3, lty=2)
    #print(p)
    dev.off()  #important to shut down the active png device
    insertImage(wb, "GRAFİK", "graph.png", width=11.18, height=7.82, units="in")
    
    
    
    
    writeData(wb, "MADDE ANALİZİ", maddeanalizi ,
              colNames = TRUE,
              rowNames = T,
              borders = "all",
              #firstActiveRow = 2,
              headerStyle = hs,
              #overwrite = F
    )
    
    setColWidths(wb, sheet = "MADDE ANALİZİ", cols = 1:3, widths = c(15,15,15))
    
    
    
    # 
    # writeData(wb, "MADDE ANALİZİB", maddeanalizib ,
    #           colNames = TRUE,
    #           rowNames = T,
    #           borders = "all",
    #           #firstActiveRow = 2,
    #           headerStyle = hs,
    #           #overwrite = F
    # )
    # 
    # setColWidths(wb, sheet = "MADDE ANALİZİB", cols = 1:3, widths = c(15,15,15))
    # 
    
    
    
    
    saveWorkbook(wb, file=choose.files(caption="KAYDETMEK İSTEDİĞİNİZ DOSYA ADINI YAZINIZ", 
                                       filters = c("Excel Dosyaları (.xlsx)","*.xlsx")), overwrite=TRUE)
    
    
  }
  
  
  
  if (iptal_edilecek_soru_kararı == 1) {
    soru_iptalleri <-
      readline(prompt = "İPTALİ İSTENEN SORULARI ARALARINA VİRGÜL KOYARAK GİRİNİZ: ")
    iptaller <- as.integer(unlist(strsplit(soru_iptalleri, ",")))
    
    # soru_iptalleriB <-
    #   readline(prompt = "B KİTAPÇIĞINDA İPTALİ İSTENEN SORULARI ARALARINA VİRGÜL KOYARAK GİRİNİZ: ")
    # iptallerB <- as.integer(unlist(strsplit(soru_iptalleriB, ",")))
    # 
    
    
    
    tamveri <- tamveri[, -(12 + c(iptaller))]
    
    
   # tam_okumaB <- tam_okumaB[, -(12 + c(iptallerB))]
    
    
    
    
    cat("İPTALLERDEN SONRA SINAVDA ",
        paste0(dim(tamveri)[2] - 12),
        " SORU KALMIŞTIR",
        "\n")
    
    
    
    ##### PUANLAMA BİLGİLERİ ###### BAKILACAK
    
    
   # puanlamabilgileriA <- tamveri[which(tamveri$KitapçıkTürü =="A"),]
   # puanlamabilgileriB <- tamveri[which(tamveri$KitapçıkTürü =="B"),]
    
    
    #puanlamabilgileri<-rbind(puanlamabilgileriA,puanlamabilgileriB)
    
    #####
    
    
    
    dy_matrisi <-
      matrix(NA,nrow= (dim(tamveri)[1]), ncol=(dim(tamveri)[2] )-12)
    
    
    # 
    # dy_matrisiB<-
    #   matrix(NA,nrow= (dim(tam_okumaB)[1]), ncol=(dim(tam_okumaB)[2] )-12)
    # 
    
    # A Kitapçığı 
    for (i in 1:dim(tamveri)[1]) {
      
      dy_matrisi[i, ] <-
        ifelse(tamveri[i, 13:dim(tamveri)[2]] == tamveri[1, 13:dim(tamveri)[2]], 1, 0)
      #Soru İptalleri için düzeltme
      
      
      
    }
    #Soru İptalleri için düzeltme
    #dy_matrisia[,c(4,7,30,34)]<-1
    
    # 
    # # B Kitapçığı
    # for (i in 1:dim(tam_okumaB)[1]) {
    #   
    #   dy_matrisiB[i, ] <-
    #     ifelse(tam_okumaB[i, 13:dim(tam_okumaB)[2]] == tam_okumaB[1, 13:dim(tam_okumaB)[2]], 1, 0)
    #   #Soru İptalleri için düzeltme
    #   
    #   
    #   
    # }
    #Soru İptalleri için düzeltme
    #dy_matrisib[,c(45,46,21,17)]<-1
    
    # dy_matrisi<-rbind(dy_matrisiA,dy_matrisiB)
    
    
    
    
    
    anahtar<- as.character(as.vector(tamveri[1,-c(1:12)]))   
    
   # anahtarb<- as.character(as.vector(tam_okumaB[2,-c(1:12)]))   
    
    
    #seceneka<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="A"),]
    
    secenek<-tamveri[-1,-c(1:12)]
    
    
    #secenekb<-tam_okuma[which(tam_okuma$`Kitapçık Türü` =="B"),]
    
    #secenekB<-tam_okumaB[-1,-c(1:12)]
    
    
    
    
    
    
    dy10 <- apply(dy_matrisi, 2, as.numeric)
    
    dsay <- apply(dy10, 1, sum)
    
    
    
    Puan <- round(dsay * tam_puan /(madde_sayisi-length(iptaller)), 0)
    
    
    
    
    
    
    ### PUAN BETİMLEME İÇİN PUAN
    
    
    dy_matrisiP<-dy_matrisi[-1,]
    
    
    
    
    
    dy10P <- apply(dy_matrisiP, 2, as.numeric)
    
    dsayP <- apply(dy10P, 1, sum)
    
    
    
    PuanP <- round(dsayP * tam_puan / madde_sayisi, 0)
    ############################################################
    
    
    yazılacakdosya0 <-
      data.frame(cbind(puanlamabilgileri[, 1:12], dsay, Puan))
    names(yazılacakdosya0)[13] <- "Doğru Cevap Sayısı"
    yazılacakdosya1 <-
      yazılacakdosya0[order(yazılacakdosya0$Puan,decreasing = T), ]
    
    ### KATILMA DURUMU DÜZENLEMELERİ   #####
    
    yazılacakdosya1$SınavaKatılmama<-"KATILDI"
    # tam_okumaKatılmayanlar$SınavaKatılmama<-"KATILMADI"
    # tam_okumaKatılmayanlar1<-tam_okumaKatılmayanlar[c(1:12)]
    # tam_okumaKatılmayanlar1$'Doğru Cevap Sayısı'<- "-"
    # tam_okumaKatılmayanlar1$Puan<- "-"
    # 
    
    ### İPTAL EDİLEN DÜZENLEMELERİ ######
    
    # 
    # tam_okumaİptalEdilenler$Sınavİptali<-"İPTAL EDİLDİ"
    # tam_okumaİptalEdilenler1<-tam_okumaİptalEdilenler[c(1:12)]
    # tam_okumaİptalEdilenler1$'Doğru Cevap Sayısı'<- "-"
    # tam_okumaİptalEdilenler1$Puan<- "-"
    # 
    # 
    
    
    
    #yazdırılacak2<-rbind(yazılacakdosya1,tam_okumaKatılmayanlar1,tam_okumaİptalEdilenler1)
    
    yazdırılacak2<-yazdırılacakdosya1
    
    ####### MADDE ANALİZİ ##############
    
    
    
    
    
    #secenek<-secenekA
    
    maddeanalizi<-ITEMAN3(key = anahtar)
    
    # 
    # secenek<-secenekB
    # 
    # maddeanalizib<-ITEMAN3(key = anahtarb)
    # 
    
    
    ### YAZDIRMA İŞLEMLERİ #####
    
    
    
    
    hs <-
      createStyle(
        textDecoration = "BOLD",
        fontColour = "#FFFFFF",
        fontSize = 12,
        fontName = "Arial Narrow",
        fgFill = "#4F80BD"
      )
    
    x <- chartr("\\", "/", veriyolu)
    
    
    dosyaadı <- tail(strsplit(x, "/")[[1]], 1)
    
    
    
    wb <- createWorkbook()
    addWorksheet(wb, "SONUÇLAR")
    addWorksheet(wb, "GRAFİK", gridLines=FALSE)
    addWorksheet(wb, "BETİMSEL", gridLines=T)
    addWorksheet(wb, "MADDE ANALİZİ", gridLines=T)
    
    
    
    
    
    
    writeData(wb, "SONUÇLAR", yazdırılacak2,
              colNames = TRUE,
              borders = "all",
              #firstActiveRow = 2,
              headerStyle = hs,
              withFilter = T
              #overwrite = F
    )
    
    
    setColWidths(wb, sheet = "SONUÇLAR", cols = 1:13, widths = "auto")
    
    freezePane(wb, "SONUÇLAR", firstRow = T)
    
    betimsel<-round(describe(PuanP),2)
    
    betimsel1<-betimsel[,c(2,3,4,5,8,9)]
    
    colnames(betimsel1)<-c("Birey_Sayısı","Ortalama_Puan","S.Sapma","Ortanca","Min","Max")
    
    
    
    
    
    writeData(wb, "BETİMSEL",betimsel1 ,
              colNames = TRUE,
              borders = "surrounding",
              #firstActiveRow = 2,
              headerStyle = hs,
              #overwrite = F
    )
    
    setColWidths(wb, sheet = "BETİMSEL", cols = 1:6, widths = "auto")
    
    
    
    
    
    
    
    png("graph.png", width=1024, height=768, units="px", res=144)  #output to png device
    p <- hist(Puan,main="ÖLÇME ve DEĞERLENDİRME MERKEZİ",col = "red",ylab = "FREKANS", xlim = c(0,100))
    abline(v=60,col="green",lwd=3, lty=2)
    #print(p)
    dev.off()  #important to shut down the active png device
    insertImage(wb, "GRAFİK", "graph.png", width=11.18, height=7.82, units="in")
    
    
    
    
    writeData(wb, "MADDE ANALİZİ", maddeanalizi ,
              colNames = TRUE,
              rowNames = T,
              borders = "all",
              #firstActiveRow = 2,
              headerStyle = hs,
              #overwrite = F
    )
    
    setColWidths(wb, sheet = "MADDE ANALİZİ", cols = 1:3, widths = c(15,15,15))
    
    
    
    # 
    # writeData(wb, "MADDE ANALİZİB", maddeanalizib ,
    #           colNames = TRUE,
    #           rowNames = T,
    #           borders = "all",
    #           #firstActiveRow = 2,
    #           headerStyle = hs,
    #           #overwrite = F
    # )
    # 
    # setColWidths(wb, sheet = "MADDE ANALİZİB", cols = 1:3, widths = c(15,15,15))
    
    
    
    
    
    saveWorkbook(wb, file=choose.files(caption="KAYDETMEK İSTEDİĞİNİZ DOSYA ADINI YAZINIZ", 
                                       filters = c("Excel Dosyaları (.xlsx)","*.xlsx")), overwrite=TRUE)
    
    
    
    
    
    
  }
  
  


}
  cat(green("İŞLEMLER BAŞARIYLA TAMAMLANMIŞTIR\n"))
  options(warn = defaultW)
  
}
