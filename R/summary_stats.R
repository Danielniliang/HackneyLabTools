#' Generate Demographic Characteristic Table
#'
#' Creates formatted descriptive summary tables for continuous and categorical variables by groups,
#' including statistical tests (ANOVA, t-test, Mann-Whitney U, Chi-square, Fisher's).
#'
#' @param con A matrix or data frame of continuous variables.
#' @param cat A matrix or data frame of categorical variables.
#' @param group A factor or vector indicating group membership.
#' @param median Logical. If TRUE, uses median/IQR and non-parametric tests; otherwise, mean/SD and parametric tests.
#' @param Fisher Logical. If TRUE, uses Fisher's exact test for categorical variables.
#' @param Ncolumn Logical. If TRUE, adds N columns for each group.
#' @param paired Logical. If TRUE, applies paired tests (e.g., paired t-test, Wilcoxon).
#' @param decimal Number of decimal places for statistics. Default is 2.
#' @param outfile File name for the Excel output. Default is 'Descriptive Analysis Result.xlsx'.
#'
#' @return A data frame of the formatted summary table, and exports to Excel.
#' @export
DemoTable <- function(con = c(),
                      cat = c(),
                      group,
                      median = FALSE,
                      Fisher = FALSE,
                      Ncolumn = FALSE,
                      paired = FALSE,
                      decimal = 2,
                      outfile = 'Descriptive Analysis Result.xlsx') {
  if(length(con)>0){
    # Tables for continuous variabels
    data<-as.data.frame(apply(apply(con,2,as.character),2,as.numeric))
    da<-as.data.frame(cbind(data,group))
    gt<-group_by(da, group)
    group<-factor(group)
    #Tables for mean ± SD, P by ANOVA
    if(median==F){
      pval<-c()
      e<-c()
      x=0
      for (i in 1:ncol(data)) {
        x=x+1
        sum <- summarize_at(gt,variable.names(data)[i],list(~mean(.,na.rm=T),~sd(.,na.rm=T)))
        mean_sd<-apply(as.matrix(sum)[,2:3], 2, as.numeric)
        #tmp<-paste(round(mean_sd[,1],2),'±', round(mean_sd[,2],1),sep=" ")
        m<-paste(format(round(mean_sd[,1],decimal),nsmall=decimal),'±', format(round(mean_sd[,2],decimal),nsmall = decimal),sep=" ")
        #m1<-c(paste(round(mean(data[,i],na.rm = TRUE),2),'±', round(sd(data[,i],na.rm = TRUE),1),sep=" "),m)
        m1<-c(paste(format(round(mean(data[,i],na.rm = TRUE),decimal), nsmall = decimal),'±', format(round(sd(data[,i],na.rm = TRUE),decimal),nsmall = decimal),sep=" "),m)

        e<-rbind(e,m1)
        #Calculate P value using t test or One way ANOVA
        x=x+1
        sum1 <- summarize_at(gt,variable.names(data)[i],funs(mean(.,na.rm=T),sd(.,na.rm=T)))
        if(is.na(sum(sum1[,2]))){pv<-NA
        }else{
          if(paired==T){
            pv<-round(t.test(data[which(group==levels(group)[1]),i],data[which(group==levels(group)[2]),i] ,paired=T)$p.value,3)
          }else{
            if(nlevels(group)>2){
              temp<-Anova(lm(data[,i]~group))
              # test<-unlist(summary(temp))
              pv<-round(temp[1,4],3)
            }else{
              Ftest<-var.test(data[,i]~group)
              if(Ftest$p.value<0.05){
                test<-t.test(as.numeric(data[,i])~group,var.equal=F)
              }else{test<-t.test(as.numeric(data[,i])~group,var.equal=T)
              }
              pv<-round(test$p.value,3)
            }
          }
        }

        if(is.na(pv)){pval<-rbind(pval,"NA")
        }else{
          if(pv<0.05& pv>=0.01){
            pval<-rbind(pval,paste(pv,'*'))
          }else{
            if(pv<0.01 & pv>=0.001){
              pval<-rbind(pval,paste(pv,'**'))
            }else{
              if(pv>=0.05){
                pval<-rbind(pval,paste(pv,""))
              }else{
                if(pv<0.001){
                  pval<-rbind(pval,"<.001 **")
                }else{
                  pval<-rbind(pval,"NA")
                }
              }
            }
          }
        }
      }
      totaln<-sum(table(group))
      np<-table(group)
      n<-c(totaln,np,"")
      e<-rbind(n,cbind(e,pval))
      rownames(e)<-c("n",variable.names(data))
      colnames(e)<-c("Total",levels(factor(group)),"p-value")
    }else{

      #Tables for Median (IQR), P by U test (For study group = 2)
      pval<-c()
      e<-c()
      x=0
      for (i in 1:ncol(data)) {
        x=x+1
        sum <- summarize_at(gt,variable.names(data)[i],list(~median(.,na.rm=T),~IQR(.,na.rm=T)))
        m<-c()
        for (j in 1:length(levels(factor(group)))) {
          tmp<-paste(round(sum[j,2],1),' (', round(sum[j,3],2),')',sep="")
          m<-c(m,tmp)
        }
        m1<-c(paste(round(median(data[,i],na.rm = TRUE),1),' (', round(IQR(data[,i],na.rm = TRUE),2),')',sep=""),m)

        e<-rbind(e,m1)
        #Calculate P value using Mann Whitney U test or Kruskal Wallis Test One Way Anova
        x=x+1
        sum1 <- summarize_at(gt,variable.names(data)[i],funs(mean(.,na.rm=T),sd(.,na.rm=T)))

        if(is.na(sum(sum1[,2]))){
          pv<-NA
        }else {
          if(nlevels(group)>2){
            pv<-round(kruskal.test(data[,i]~group)$p.value ,3)
          }else{
            if(paired==T){
              pv<-round(wilcox.test(data[which(group==levels(group)[1]),i],data[which(group==levels(group)[2]),i] ,paired=T)$p.value,3)
            }else{
              pv<-round(wilcox.test(data[,i]~group)$p.value,3)
            }

          }
        }

        if(is.na(pv)){pval<-rbind(pval,"NA")
        }else{
          if(pv<0.05& pv>=0.01){
            pval<-rbind(pval,paste(pv,'*'))
          }else{
            if(pv<0.01 & pv>=0.001){
              pval<-rbind(pval,paste(pv,'**'))
            }else{
              if(pv>=0.05){
                pval<-rbind(pval,paste(pv,""))
              }else{
                if(pv<0.001){
                  pval<-rbind(pval,"<.001 **")
                }else{
                  pval<-rbind(pval,"NA")
                }
              }
            }
          }
        }
      }

      totaln<-sum(table(group))
      pct<-round(prop.table(table(group))*100,2)
      np<-paste(table(group),' (',pct,')', sep='')
      n<-c(totaln,np,"")

      e<-rbind(n,cbind(e,pval))
      rownames(e)<-c("n",variable.names(data))
      colnames(e)<-c("Total",levels(factor(group)),"p-value")
    }
  }

  if(length(cat)>0){
    #Tables for categorical variabes
    data1<-as.data.frame(apply(cat,2,factor))
    f<-c()
    x=0
    for (i in 1:ncol(data1)) {
      if(length(data1[,i][which(is.na(data1[,i])==T)])==length(data1[,i])){next}
      x=x+1
      m <-t(rbind(table(data1[,i]),table(group,data1[,i])))
      #rownames(m)<-c(variable.names(data1)[i],rep(" ",length(levels(factor(data1[,i])))))
      n<-table(group,data1[,i])
      cols<-c()
      for (j in 1:ncol(m)) {
        cols<-c(cols,sum(m[,j]))
      }
      #Formating and Calculating P values
      for (j in 1:nrow(m)){
        for (k in 1:ncol(m)){
          m[j,k] = paste(m[j,k],' (',round(as.numeric(m[j,k])*100/cols[k],1),')',sep = '')
        }
        if(Fisher==F){
          pv<-chisq.test(n)
        }else{
          pv<-fisher.test(n)
        }
        if(pv$p.value<0.001){
          pval<-"<.001 **"
        }else{
          if(pv$p.value<0.01){
            pval<-paste(round(pv$p.value,3),"**")
          }else{
            if(pv$p.value<0.05& pv$p.value>=0.01){
              pval<-paste(round(pv$p.value,3),"*")
            }else{
              pval<-round(pv$p.value,3)
            }
          }
        }
      }
      m<-cbind(rownames(m),m,"")
      m<-rbind(c(rep("",ncol(m)-1),pval),m)


      rownames(m)<-c(variable.names(data1)[i],rep(" ",nrow(m)-1))
      f<-rbind(f, "", m)
    }
    colnames(f)<-c("","Total",levels(factor(group)),"")
  }



  #option of adding N columns for each group:
  #Function of adding N columns:
  if(Ncolumn==T){
    Ntable<-function(vi,result,group){
      n<-c()
      for (i in 1:ncol(vi)) {
        tmp<-cbind.data.frame(is.na(vi[,i]),group)
        tmp1<-table(tmp[,1],tmp[,2])
        n<-rbind(n,tmp1[1,])
      }
      rownames(n)<-colnames(vi)

      nresult<-c()
      for (i in 1:ncol(n)) {
        temp<-cbind(c("",n[,i]),result[,i+1])
        nresult<-cbind(nresult,temp)
      }
      nresult<-cbind(result[,1],nresult,result[,ncol(result)])
      level<-c()
      for (i in 1:nlevels(factor(group))) {
        level<-cbind(level,cbind('N',levels(factor(group))[i]))
      }
      colnames(nresult)<-c('Total',level,"P value")
      return(nresult)
    }
    e<-Ntable(vi=con,result = e,group = group)
  }


  if(length(con)>0&length(cat)==0){
    table<-e
  }else{
    if(length(cat)>0&length(con)==0){
      table<-f
    }else{
      if(Ncolumn==T){

        ftmp=c()
        for (i in 2:(ncol(f)-2)) {
          ftmp=cbind(ftmp,cbind(f[,i],""))
        }
        f=cbind(f[,1],ftmp,f[,(ncol(f)-1):ncol(f)])
        table<-rbind(cbind("",e),f)
      }else{
        table<-rbind(cbind("",e),f)
      }
    }
  }


  y<-table
  #Notation
  if(is.null(con)==F & is.null(cat)==F){
    table<-rbind(cbind(rownames(table),table),
                 c(
                   paste("Table 1. Values are presented as",
                         ifelse(median,"Median (IQR)","Mean ± SD"),
                         "for continuous variables, and n (%) for categorical variables. P values were calculated with" ,
                         ifelse(median,
                                ifelse(nlevels(group)!=2, "Kruskal Wallis Test One Way ANOVA","Mann-Whitney U test"),
                                ifelse(nlevels(group)!=2, "One-way ANOVA",ifelse(paired,"Paired t test","independent t test"))),
                         "for continuous variables and",
                         ifelse(Fisher, "Fisher's exact test", "Chi-square test"),
                         "for categorical variables."),
                   rep("",ncol(table))
                 ))
  }else{
    if(is.null(con)==T & is.null(cat)==F){
      table<-rbind(cbind(rownames(table),table),
                   c(
                     paste("Table 1. Values are presented as n (%). P values were calculated with" ,
                           ifelse(Fisher, "Fisher's exact test", "Chi-square test")),
                     rep("",ncol(table))
                   ))
    }else{
      if(is.null(con)==F & is.null(cat)==T){
        table<-rbind(cbind(rownames(table),table),
                     c(
                       paste("Table 1. Values are presented as",
                             ifelse(median,"Median (IQR).","Mean ± SD."),
                             "P values were calculated with" ,
                             ifelse(median,
                                    ifelse(nlevels(group)!=2, "Kruskal Wallis Test One Way ANOVA","Mann-Whitney U test"),
                                    ifelse(nlevels(group)!=2, "One-way ANOVA",ifelse(paired,"Paired t test","independent t test")))),
                       rep("",ncol(table))
                     ))
      }
    }
  }



  # export(as.data.frame(table),outfile)
  ## Create a workbook
  x<-rbind(colnames(table),table)
  colnames(x)<-rep(paste("Table 1. Demographic Characteristics",levels(group)[1],"vs.",levels(group)[2]),ncol(x))
  #write and designate a workbook
  wb <- write.xlsx(x, outfile)

  #Customize style:
  First2<-createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin',textDecoration = "bold")
  Border <- createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin')
  bold<-createStyle(textDecoration = "bold")
  wraptext<-createStyle(wrapText = T,valign = 'top',halign = 'left')
  #Apply style:
  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = T)
  if(Ncolumn==T){
    addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = T)
  }else{
    addStyle(wb, sheet = 1, First2, rows = 2, cols = 3:ncol(x), gridExpand = T)
  }
  #addStyle(wb, sheet = 1, Border, rows = nrow(x)-1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, Border, rows = nrow(x), cols = 1:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, bold, rows = 3:nrow(x)-3, cols = 1, gridExpand = T)

  ## Merge cells: Merge the footer's rows
  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  #mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x))
  mergeCells(wb, 1, cols = 1:ncol(x), rows = (nrow(x)+1):(nrow(x)+4))
  addStyle(wb,1,wraptext,cols = 1:ncol(x), rows = nrow(x)+1)

  #set colunme width:
  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto",ignoreMergedCells = T)
  #setRowHeights(wb, 1, rows = nrow(x)+1, heights  = "auto")

  #Modify font:
  modifyBaseFont(wb, fontSize = 12, fontColour = "black",fontName = "Times New Roman")

  ## Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)

  openXL(outfile)
  return(y)
}

#' Perform Univariate Group Comparison
#'
#' Conducts statistical tests to compare continuous outcomes between two groups,
#' providing p-values, effect sizes (Cohen's d), and summary statistics (mean/SD or median/IQR).
#'
#' @param dt A data frame or matrix of numeric variables to compare.
#' @param group A vector or factor indicating the grouping variable.
#' @param median Logical. If TRUE, uses non-parametric tests and reports medians.
#' @param paired Logical. If TRUE, uses paired tests.
#' @param cohend.correct Logical. If TRUE, applies Hedges' correction to Cohen's d.
#' @param outfile Output Excel file name. Default is 'Comparison of difference between group result.xlsx'.
#'
#' @return A data frame of group comparison results, and exports to Excel.
#' @export
group.diff <- function(dt, group,
                       median = FALSE,
                       paired = FALSE,
                       cohend.correct = TRUE,
                       outfile = 'Comparison of difference between group result.xlsx') {
  #Transform the data into numeric
  dt<-as.data.frame(apply(dt,2,as.numeric))

  #calculate the mean and standard deviation for each group
  da<-group_by(cbind.data.frame(dt,group),group)
  result1<-as.data.frame(summarise_all(da,funs(mean(.,na.rm=T),sd(.,na.rm=T),median(.,na.rm = T),IQR(.,na.rm = T)))[,-1])
  rownames(result1)<-levels(group)
  mean<-result1[,1:ncol(dt)]
  sd<-result1[,(ncol(dt)+1):(2*ncol(dt)),]
  Median<-result1[,(2*ncol(dt)+1):(3*ncol(dt)),]
  iqr<-result1[,(3*ncol(dt)+1):(4*ncol(dt)),]
  mean_dff<-unlist(mean[2,]-mean[1,])

  #Testing:
  #Testing for continues Var

  result2<-c()
  for (i in 1:ncol(dt)) {
    # N missing:
    N_missing<-sum(is.na(dt[,i]))
    if(median==F){

      #t test:
      if(paired==T){
        group<-factor(group)
        test<-t.test(dt[which(group==levels(group)[1]),i],dt[which(group==levels(group)[2]),i] ,paired=T)
      }else{
        Ftest<-var.test(dt[,i]~group)
        if(Ftest$p.value<0.05){
          test<-t.test(as.numeric(dt[,i])~group)
        }else{test<-t.test(as.numeric(dt[,i])~group,var.equal=TRUE)
        }
      }

    }else{
      #U test:
      if(paired==T){
        test<-wilcox.test(dt[which(group==levels(group)[1]),i],dt[which(group==levels(group)[2]),i] ,paired=T)
      }else{
        test<-wilcox.test(as.numeric(dt[,i])~group)
      }

    }

    groupi<-relevel(group,ref = levels(group)[2])
    #Cohen's d
    if(cohend.correct==T){
      if(paired==F){
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=FALSE,
                       na.rm=FALSE, hedges.correction=TRUE,
                       conf.level=0.95,noncentral=FALSE)
      }else{
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=TRUE,
                       na.rm=FALSE, hedges.correction=TRUE,
                       conf.level=0.95,noncentral=FALSE)
      }
    }else{
      if(paired==F){
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=FALSE,
                       na.rm=FALSE, hedges.correction=FALSE,
                       conf.level=0.95,noncentral=FALSE)
      }else{
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=TRUE,
                       na.rm=FALSE, hedges.correction=FALSE,
                       conf.level=0.95,noncentral=FALSE)
      }
    }

    #Formatting:
    cd<-paste(round(cohen$estimate,2)," (",round(cohen$conf.int[1],1), " ,",round(cohen$conf.int[2],1),")",sep = "")



    if(median==F){
      meandiff<-paste(round(mean_dff[i],2)," (",round(test$conf.int[1],2), " ,",round(test$conf.int[2],2),")",sep = "")
      result2<-rbind(result2, c(meandiff, round(test$p.value,3), cd, N_missing))
    }else{
      result2<-rbind(result2, c(round(test$p.value,3), cd, N_missing))
    }
  }

  if(median==F){
    result<-cbind(round(t(mean),2),result2)
    result[,1:2]<-cbind(paste(result[,1], "±",round(sd[1,],2)),paste(result[,2], "±",round(sd[2,],2)))
    rownames(result)<-colnames(dt)
    #colnames(result)<-c(paste(levels(group)[1],"(Mean ± SD)"),paste(levels(group)[2],"(Mean ± SD)"),"Mean Difference",
    #                    "CI Lower","CI Upper","statistics","P-value","cohen's D","CI Lower","CI Upper","Missing")
  }else{
    result<-cbind(round(t(Median),2),result2)
    result[,1:2]<-cbind(paste(result[,1], " (",round(iqr[1,],2),")",sep = ""),paste(result[,2], " (",round(iqr[2,],2),")",sep = ""))
    rownames(result)<-colnames(dt)
    #colnames(result)<-c(paste(levels(group)[1],"(Median (IQR))"),paste(levels(group)[2],"(Median (IQR))"),
    #                    "statistics","P-value","cohen's D","CI Lower","CI Upper","Missing")
  }
  #N value for result:
  #N:
  vi<-dt
  n<-c()
  for (i in 1:ncol(vi)) {
    tmp<-cbind.data.frame(is.na(vi[,i]),group)
    tmp1<-table(tmp$`is.na(vi[, i])`,tmp$group )
    tmp2<-sum(is.na(vi[,i])==F)
    n<-rbind(n,c(tmp2,tmp1[1,]))
  }
  result4<-cbind(n[,2], result[,1],n[,3],result[,2],result[,3:ncol(result)],n[,1])

  #Naming columns:
  if(median==F){
    colnames(result4)<-c("N",paste(levels(group)[1],"(Mean ± SD)"),"N",paste(levels(group)[2],"(Mean ± SD)"),"Mean Difference",
                         "P-value","Cohen's D (CI)","N Missing","Total N")
  }else{
    colnames(result4)<-c("N",paste(levels(group)[1],"(Median (IQR))"),"N",paste(levels(group)[2],"(Median (IQR))"),
                         "P-value","Cohen's D (CI)","N Missing","Total N")
  }

  y<-result4
  ####Formating worksheet#####
  #Title:
  result4<-cbind(rownames(result4),result4)
  x<-rbind(colnames(result4),result4)
  colnames(x)<-rep(paste("Table 2. ",levels(group)[1],"vs.",levels(group)[2]),ncol(x))

  #Footnote:
  if(median==F){
    x<-rbind(x,
             "Comparison of outcome variables between. P values were obtained by independent t-test. ")
  }else{
    x<-rbind(x,
             "Comparison of outcome variables between. P values were obtained by Mann-Whitney U test")
  }


  #write and designate a workbook
  wb<-createWorkbook()
  addWorksheet(wb, "Analysis Results")
  writeData(wb,1,x)
  #wb <- write.xlsx(x, outfile)
  #Customize style:
  First2<-createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin',textDecoration = "bold")
  Border <- createStyle(border="TopBottom", borderColour = "black",borderStyle = 'thin')
  bold<-createStyle(textDecoration = "bold")
  #Apply style:
  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, Border, rows = nrow(x)-1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, Border, rows = nrow(x)+1, cols = 1:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, bold, rows = 3:nrow(x)-3, cols = 1, gridExpand = T)

  ## Merge cells: Merge the footer's rows
  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  #mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x))
  mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x)+1)

  #set colunme width:
  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto",ignoreMergedCells = T)

  #Modify font:
  modifyBaseFont(wb, fontSize = 12, fontColour = "black",fontName = "Times New Roman")

  ## Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)

  openXL(outfile)

  return(y)
}
