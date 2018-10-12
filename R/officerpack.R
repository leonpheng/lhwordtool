#' fnote
#'
#' convert and sorte list of define in footnote to one paragraph.
#' @param data list of define. ex<-c("A","C","B"); fnote(ex)= "ABC"
#'
#' @keywords fnote
#' @export
#' @examples
#' fnote()
#'
fnote<-function(data){
  data<-paste(sort(data), sep = '', collapse = '')
  data}
#' len
#'
#' could be used in lhtext
#' @param t ex: t[[1]] and the subsequent t is t[[len(t)]]. Convenience is you can insert t after the first t
#'
#' @keywords len
#' @export
#' @examples
#' len()
#'
len<-function(t){
  x<-length(t)+1
  x
}



#' officer.report.template
#'
#' Create doc for word document using Officer.
#' 
#' @param TFL FAULT for no table and figure lists
#'
#' @keywords officer.report.template
#' @export
#' @examples

officer.report.template<-function (temp="c:/lhtemplate/stylereport.docx",TFL=T)
{

  library(officer)
  library(flextable)
  library(magrittr)
  doc<-read_docx(temp)
   
   if (TFL) {
     doc <- body_add_break(doc)
  
   } 
  doc
  }



#' lhtext
#'
#' Simple way to create word document using loop of from t list.
#' Type lhtext and copy the template to R workspace and start writing.
#' @param t list of items. see example.
#' @param toc.level maximimum toc level
#' @param template Word document template could be used for styles. Styles should be mapped in style.to.map. Template is also available at github: to load it, just run  lhtemp() once to download and store the templates in your PC at "c:lhtemplate. Note that the templates and logo are also used in xptdef package.
#' @param TOC Set to F if no TOC wanted
#' @param style.to.map Map the styles in template to be used. Ex: mypar is for footnote (font size)
#' @keywords lhtext
#' @export
#' @examples
#' TEMPLATE
#' x<-NULL
#' t[[len(t)]]<-c("lev1","Ceci est un test. lev1= level 1 title")
#' t[[len(t)]]<-c("txtn","bold","Add bold paragragh with page break","pageb") #normal simple text
#' t[[len(t)]]<-c("lev2","Title for level 2. Below is for creating table")
#' tab0<-data.frame(x=c("MONDAY","TUESDAY"),y=c("EAT SOMETHING","DON'T EAT"))
#' t[[len(t)]]<-c("tab","dataframe")
#' t[[len(t)]]<-c("fnt","This is foot not with font size of 9 by default","pageb")
#' t[[len(t)]]<-c("lev2","This a test Plot- Used plot. ggplot could be inserted")
#' t[[len(t)]]<-c("fig","plot(1,1)",5,5,"pageb")
#' tab1<-read.csv("1-OH-Midazolam1.25+1.25mg.csv")
#' t[[len(t)]]<-c("lev2","Another Table from folder")
#' t[[len(t)]]<-c("tab","FlexTable(tab1)","pageb")
#' t[[len(t)]]<-c("lev2","This a test Image imported from folder and inserted in doc")
#' t[[len(t)]]<-c("fcap" or "tcap","This is for figure caption and tcap for table caption")
#' t[[len(t)]]<-c("ima","hydroxy-auc0.6-v-age-1.png",7.4,6)
#' t[[len(t)]]<-c("txtc","This is complex text editor is ::b:i"," X","2::e:u" ,"bold::b:i:s","fwb::e","this::i","fsti::b:i") #formatted text.
#' tips for symbol, example: ++a++, ++inf++, ++n++, etc.. then do search and replace in Word
#' e=superscrip, s=subscript, b=bold, i=italic, ::= start code, := additional code
#' lhtext(t,toc.level=3,template="style.docx",TOC=F)
#' to write doc: writeDoc(doc,"name.docx")
#' lhtext()

lhtext<-function (t)
{
  library(ReporteRs)
  library(flextable)
  library(dplyr)
  library(plyr)
library(stringr)
  library(officer)  
for (i in 1:length(t)) {
    b <- function(x) {}
    if (substring(t[[i]][1], 1, 3) == "lev") {
      l <- gsub("lev", "", t[[i]][1])
      doc<-body_add_par(doc,value =  t[[i]][2],style =paste0("HD",l))
      #doc <- addTitle(doc, t[[i]][2], level = as.numeric(l))
    }
#TABLE CAPTION    
  if (t[[i]][1] == "tcap") {
      doc<-body_add_par(doc,value = t[[i]][2],style ="tabcaption")}
#TITLE  
  if (t[[i]][1] == "title") {
      doc<-body_add_par(doc,value = t[[i]][2],style ="lhtitle") } 
      
  #    doc <-addParagraph(doc,value = t[[i]][2],
  #                       stylename = "tabcaption")}

#FIGURE CAPTION    
  if (t[[i]][1] == "fcap") {
        doc<-body_add_par(doc,value = t[[i]][2],style ="figcaption") }

    if (t[[i]][1] == "pgb"){
      doc <- body_add_break(doc)
    }
#PRESET FOOTNOTE (use txt9 with more functionality)   
  if (t[[i]][1] == "fnt") {
    if(t[[i]][3]=="pgb"){
      doc <-  body_add_par(doc, t[[i]][2], style = "fnt")
      doc <- body_add_break(doc)}else{
        doc <-  body_add_par(doc, t[[i]][2], style = "fnt")
      }
  }

#FIGURE    
  if (t[[i]][1] == "fig") {
    b <- function(x) {}
      body(b) <- parse(text = t[[i]][2])
      doc <- body_add_gg(doc,value = t[[i]][2], style = "center" ) 
    #  doc <- addPlot(doc, fun = function() b(), width = as.numeric(t[[i]][3]),
   #                  height = as.numeric(t[[i]][4]), par.properties = parProperties(text.align = "center"))
    }
    if (t[[i]][1] == "ima") {
      doc <-body_add_img(doc,src = t[[i]][2], width = as.numeric(t[[i]][3]), height =  as.numeric(t[[i]][4]), style = "center")
    }

#TABLE PRESET
    def_cell <- fp_cell(border = fp_border(color="black"))
    std_b <- fp_border(color="black")
    def_par <- fp_par(text.align = "center")
    def_text <- fp_text(color="black", italic = F,font.family="Time Roman")
    def_text_header <- update(color="black", def_text, bold = TRUE)    
#TABLE  
    if (t[[i]][1] == "tab") {
     b <- function(x) {}
     body(b) <- parse(text = t[[i]][2])
     if(is.data.frame(b())){
     ft<-regulartable(b())
     ft <- style( ft, pr_c = def_cell, pr_p = def_par, pr_t = def_text, part = "all")
     ft <- style( ft, pr_t = def_text_header, part = "header")
     ft<-bg(ft,bg="grey",part="header")
     ft <- theme_booktabs(ft)
     ft <- border_outer( ft, border = std_b, part = "all" )
     ft <- border_inner_h( ft, border = std_b, part = "all" )
     ft <- border_inner_v( ft, border = std_b, part = "all" )
     ft <- align( ft, align = "center", part = "all" )
     ft <- autofit(ft)
doc <- body_add_flextable(doc, ft)}else{
  body(b) <- parse(text = t[[i]][2])
  ft<-b()
doc <- body_add_flextable(doc,ft)}  
    }
    
if (length(grep("txt",t[[i]][1]))==1) {
      c = t[[i]]
      all <- ""
      value <-NULL
      prop=NULL
      fs=as.numeric(gsub("txt","",t[[i]][1]))
      if(is.na(fs)){fs=12}else{fs=fs}
      if(c[length(c)]%in%c("center","left","right","justified")){
      lenc<-length(c)-1}else{lenc<-length(c)}
      for (j in 2:lenc){
        pr <- shortcuts$fp_bold(font.size = fs)
        pr <- update(pr, font.family ="Times New Roman")
        pr <- update(pr, bold =F)
        
        if (length(grep(":i", sub(".*:i", ":i", c[j]))) !=
            0) {
          pr <- update(pr, italic =TRUE)}
       

if (length(grep(":b", sub(".*:b", ":b", c[j]))) !=
            0) {
  pr <- update(pr, bold =TRUE)
        }
  if(length(grep(":s", sub(".*:s", ":s",
                                          c[j]))) != 0){
    pr <- update(pr, vertical.align	 ="subscript")
  }
    
  if(length(grep(":e", sub(".*:e", ":e",c[j]))) != 0){
    pr <- update(pr, vertical.align	 ="superscript")
  }

        if (length(grep(":u", sub(".*:u", ":u", c[j]))) !=
            0) {
          pr <- update(pr, underlined	 =TRUE)
        }
        
  
if (length(grep(":col", sub(".*:col", ":col", c[j]))) !=
    0) {
  z5 = sub(":.*","",sub(".*:col", "", c[j]))
  pr <- update(pr, color	 =z5)
}

if (length(grep(":size", sub(".*:size", ":size", c[j]))) !=
    0) {
  z6 = as.numeric(sub(":.*","",sub(".*:size", "", c[j])))
  pr <- update(pr, font.size	 =z6)
}

    if (length(grep("::", sub(".*::", "::", c[j]))) ==
        0) {
          c1 <- c[j]
      }
    else {
          c1 <- gsub(sub(".*::", "::", c[j]), "", c[j])
    }
value[[j-1]]<-c1
prop[[j-1]]<-pr
      }
for(z in 1:length(prop)){
   if(z==1){
x<-paste0("fpar(ftext(value[[",z,"]],prop =prop[[",z,"]])")}else{
x<-paste0(x,",ftext(value[[",z,"]],prop =prop[[",z,"]])")
}}
  x<-paste0(x,")") 
  b <- function(x) {
  }
body(b) <- parse(text = x)
test<-b()
if(c[length(c)]%in%c("center","left","right","justified")){
  doc <- body_add_fpar(doc,test,style=c[length(c)])}else{
    doc <- body_add_fpar(doc,test)}
#print(doc, target = "body_add_fpar_1.docx" )
      }

  }
  doc
}

####TABLE PRESET
#' lflextab
#'
#' Create doc for word document using Officer.
#' @param csv If source = csv otherwise flextable tab
#' @param lst List of header example lst=c(Mean="mean",animal="dog")
#' @param add.h Define additional header rows Ex: df<-data.frame(row1=c("",rep("median (CV%)",4),row2=c("Inches","Inches","Inches"#',"Inches","Species") unit=c("mg/mL","inch"," ",       " "," "))) then add.h=df
#' @param cf conditional formatting. should be list("i=x, j=y :(format abbreviation) format: col=color (:colred),ita=italic (:ita),bol=bold (:bol), bg=background (:bgred).  Conditional statement = i=~colname >or< or == values, j=~col1+col2". mv=vertical (merge identical value), mh=horizontal (merge identical value); ma=at (merge cells regardless values). ex. mv=list("j=1 or j=~colname :mv"); ma = c.
#' @param border Border list("vi:dashed:black:header","vo:dashed:black:body","ho:dashed:black:body",etc.)
#' @keywords officer.report.template
#' @export
#' @examples
#'
#'        

lhflextab<-function(table1,csv="yes",
                bord="yes",
                select=NULL,
                add.h= NULL,
                merge.all="yes",
                size=12,
                empty=NULL,
                cf=NULL,
                border=NULL,
                align="center"
                )
  {
b <- function(x) {}
  def_cell <- fp_cell(border = fp_border(color="black"))
  std_b <- fp_border(color="black")
  def_par <- fp_par(text.align = "center")
  def_text <- fp_text(color="black", italic = F,font.family="Time New Roman")
  def_text_header <- update(color="black", def_text, bold = TRUE)
if(!is.null(csv)){
  if(!is.null(select)){
  tab1<-regulartable(table1,col_keys=select)}else{tab1<-table1}}
  if(!is.null(empty)){ 
    for(i in 1:ncol(table1)){
      table1[,i][table1[,i]==""|is.na(table1[,i])]<-empty
      table1}
    }else{table1}  

tab1 <- style(tab1, pr_t = def_text_header, part = "header")


#For header
if(!is.null(add.h)){
    if(!is.null(select)){
    typology <-add.h}else{typology <-names(tab)}
    typology$col_keys<-select
    typology<-chclass(typology,names(typology),"char")
    tab1<-set_header_df(tab1, mapping = typology, key = "col_keys" )
    tab1 <- merge_h(tab1, part = "header")
    tab1 <- merge_v(tab1, part = "header")
    }
  
  tab1 <- style(tab1, pr_p = def_par, pr_t = def_text, part = "all")
  tab1<-bg(tab1,bg="gray88",part="header")
  tab1 <- style( tab1, pr_t = def_text_header, part = "header")
  tab1 <- fontsize(tab1,size = size, part = "all")
  std_b2 <- fp_border(color="black", style = "solid")
  std_b3 <- fp_border(color="black", style = "dashed")

if(!is.null(cf)){
  for(xx in 1:length(cf)){
    coord<-gsub(sub(".*:", ":", cf[xx]),"", cf[xx])
    fm<-gsub(sub(":.*", "", cf[xx]),"", cf[xx])
    fm<-gsub(sub(":.*", ":", fm),"", fm)
    
    if(length(grep("col",fm))==1){
      vv<-gsub("col", "", fm)
      body(b) <- parse(text =paste("color(tab1,",coord,",color=vv)"))
      tab1<- b()}     
 
    if(length(grep("mv",fm))==1){
        vv<-gsub("mv", "", fm)
        body(b) <- parse(text =paste("merge_v(tab1,",coord,")"))
        tab1<- b()  
      }  
   if(length(grep("bg",fm))==1){
      vv<-gsub("bg", "", fm)
      body(b) <- parse(text =paste("bg(tab1,",coord,",bg=vv)"))
      tab1<- b()  
   }
  if(length(grep("mh",fm))==1){
      vv<-gsub("mh", "", fm)
      body(b) <- parse(text =paste("merge_h(tab1,",coord,")"))
      tab1<- b()  
  }
    if(length(grep("ma",fm))==1){
      vv<-gsub("ma", "", fm)
      body(b) <- parse(text =paste("merge_at(tab1,",coord,")"))
      tab1<- b()
    }
    if(length(grep("bol",fm))==1){
      vv<-gsub("bol", "", fm)
      body(b) <- parse(text = paste("bold(tab1,",coord,",bold=TRUE)"))
      tab1<- b()
    }    
    if(length(grep("ita",fm))==1){
      vv<-gsub("ita", "", fm)
      body(b) <- parse(text =paste("italic(tab1,",coord,")"))
      tab1<- b()
    }
  }}

para<-fp_border(color="black", style = "dashed")
para1<-fp_border(color="black", style = "solid")
  tab1 <- border_remove(tab1)
  tab1 <- border_outer( tab1, border = para1, part = "all" )
  tab1 <- border_inner_h( tab1,border = para1,part = "all" )
  tab1 <- border_inner_v(tab1, border = para1, part = "all" )
  
if(!is.null(border)){
for(i in 1:length(border)){
  ca<-gsub(sub(":.*", ":", border[i]),"", border[i])
  co1<-gsub(ca,"",border[i]);co1<-gsub(":","",co1)
  ca1<-gsub(sub(":.*", ":", ca),"",ca)
  co2<-gsub(ca1,"",ca);co2<-gsub(":","",co2)
  ca2<-gsub(sub(":.*", ":", ca1),"",ca1)
  co3<-gsub(ca2,"",ca1);co3<-gsub(":","",co3)
  ca3<-gsub(sub(":.*", ":", ca2),"",ca2)
  co4<-gsub(ca3,"",ca2);co4<-gsub(":","",co4)

if(length(grep("out",co1))==1){
  out<-fp_border(color=co3, style = co2)
  tab1 <- border_outer( tab1, border = out, part = co4)
}
if(length(grep("vi",co1))==1){
  out<-fp_border(color=co3, style = co2)
  tab1 <- border_inner_v(tab1, border = out, part = co4)
}
if(length(grep("hi",co1))==1){
  out<-fp_border(color=co3, style = co2)
  tab1 <- border_inner_h(tab1, border = out, part = co4)
}}
 
tab1 <- align( tab1, align = align, part = "all" )  

}
tab1 <- autofit(tab1)
  }

####General functions
#' officer.report.template
#'
#' text to function.
#' @param txt If source = csv otherwise flextable tab
gfun<-function(txt){
  b <- function(x) {}
  body(b) <- parse(text =txt)
  z<-b()
  z
}

#' officer.report.template
#'
#' flxt to doc.
#' @param txt If source = csv otherwise flextable tab
flxdoc<-function(tab){
doc <- body_add_flextable (doc,tab)
}


#test purpose
# hd<-data.frame(col=c("",rep("median (CV%)",4)),col1=c("Inches","Inches","Inches","Inches","Species"),unit=c("mg/mL","inch"," "," "," "))
#   
# EX<-ah.ft(tab=dd1,
#           csv="yes",
#    bord="yes",
#     select= c("N" ,"var","mean","min","max"),
#       add.h= hd,
#     ma="1:1-1:3")
