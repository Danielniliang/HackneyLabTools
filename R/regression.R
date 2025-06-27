#' Multivariate Regression Table
#'
#' Generate a table of multivariate analysis results using linear or logistic regression.
#'
#' @param dt A data frame including all variables.
#' @param group A vector or a variable name in `dt` specifying treatment groups.
#' @param vi_name A character vector of outcome variable names.
#' @param con_name A character vector of continuous covariate names.
#' @param cat_name A character vector of categorical covariate names.
#' @param ref Reference group level name (optional).
#' @param method Regression method: "Linear" or "logistic".
#' @param CI Logical; whether to include confidence intervals in results.
#' @param log.transform Logical; whether to log-transform the outcome.
#' @param outfile Output Excel file path.
#'
#' @return A list with formatted result tables and regression summaries.
#' @export
multivariate <- function(dt,
                         group,
                         vi_name,
                         con_name,
                         cat_name,
                         ref = c(),
                         method = "Linear",
                         CI = FALSE,
                         log.transform = FALSE,
                         outfile = "Multivariate Analysis Results.xlsx") {
  dt<-as.data.frame(dt)
  vi<-dplyr::select(dt,c(vi_name))
  cov<-dplyr::select(dt,con_name,cat_name)
  cov_names<-colnames(cov)

  if(length(group)>1){
    group<-group
  }else{
    group<-dplyr::select(dt,group)
  }

  if(length(ref)!=0){
    group<-relevel(factor(group),ref = ref)
  }else{
    group=factor(group)
  }

  #Create dataset for modeling:
  if(length(con_name)>0&length(cat_name)>0){
    con_cov<-as.data.frame(apply(apply(dplyr::select(dt,con_name),2,as.character),2,as.numeric))
    cat_cov<-as.data.frame(apply(dplyr::select(dt,cat_name), 2, as.factor))
    data<-cbind.data.frame(con_cov,cat_cov,vi,group=group)
  }else{
    if(length(con_name)==0&length(cat_name)>0){
      cat_cov<-as.data.frame(apply(dplyr::select(dt,cat_name), 2, as.factor))
      data<-cbind.data.frame(cat_cov,vi,group=group)
    }else{
      if(length(con_name)>0&length(cat_name)==0){
        con_cov<-as.data.frame(apply(apply(dplyr::select(dt,con_name),2,as.character),2,as.numeric))
        data<-cbind.data.frame(con_cov,vi,group=group)
      }else{
        data<-cbind.data.frame(vi,group=group)
      }
    }
  }

  #Create Formula for model:
  if(length(cov_names)>0){
    form<-formula(paste('y','~','group','+',paste(cov_names,collapse = " + ")))
  }else{
    form<-formula(paste('y','~','group'))
  }



  re<-c()
  if(method=="logistic"){
    re2<-c()
    for (i in 1:ncol(vi)){
      y<-factor(vi[,i])
      tmp<-glm(formula = form,data = data,
               family = "binomial")
      coef<-round(summary(tmp)$coefficients,3)
      coef<-cbind.data.frame(c(vi_name[i],rep("",nrow(coef)-1)),
                             rownames(coef),coef)
      k<- exp(cbind(OR = coef(tmp), confint(tmp)))
      wald<-wald.test(b = coef(tmp), Sigma = vcov(tmp), Terms = 2:nlevels(group))
      or<-c("Ref",
            paste(round(k[2:nlevels(group),1],2),
                  " (",
                  round(k[2:nlevels(group),2],2),
                  ", ",
                  round(k[2:nlevels(group),3],2),
                  ")",sep = ""),
            round(wald$result$chi2[3],3))
      re<-rbind.data.frame(re,'',coef)
      nper<-c()
      for (j in 1:nlevels(group)) {
        ftb<-paste(table(y,group)[,j],
                   " (",
                   100*(round(prop.table(table(y,group),2),2)[,j]),
                   ")",sep = "")
        nper<-cbind(nper,ftb)
      }

      OR<-cbind.data.frame(c(vi_name[i],rep("",nlevels(y)-1)),
                           levels(y),
                           nper,
                           rbind(or,matrix("",nrow = nlevels(y)-1,ncol = nlevels(group)+1)))
      re2<-rbind(re2,'',OR)
    }
    re2<-as.matrix(re2)
    rownames(re2)<-as.matrix(re2[,1])
    re2<-re2[,-1]
    colnames(re2)<-c("",
                     levels(group),
                     paste("OR",levels(group), "vs.",ref, "(95% CI)"),
                     "P value")
    colnames(re)<-c("","",colnames(summary(tmp)$coefficients))
    #Add title and footnotes for the  result table:

    l<-list(re2,re)
    #Formating:
    wb<-ExcelTable(table = re2,
                   title ='Table 2. Multivariate Results',
                   caption = paste("Table 2. Multivariate Results. Odds ratios were obtained obtained by logistic regression model adjusting for",paste(cov_names, collapse = ', '),
                                   ". P value were obtained from Wald's test. Frequency were presented as N (%)"),
                   outfile = outfile,
                   openfile = F)
    addWorksheet(wb,"Coefficients & ANOVA")
    writeData(wb,sheet = "Coefficients & ANOVA",re)
    saveWorkbook(wb,outfile,overwrite = T)
    openXL(outfile)
  }else{
    #####For linear regression MODEL####
    re1<-c()
    betas<-list()
    for (i in 1:ncol(vi)) {
      vi[,i]<-as.numeric(as.character(vi[,i]))
      if(log.transform==T){
        y<-log(vi[,i]+1)
      }else{
        y<-vi[,i]
      }
      tmp<-lm(formula = form,data)
      #Coefficients:
      a<-as.matrix(round(coefficients(summary(tmp)),3))
      if(log.transform==T){
        a[,1:2]<-exp(a[,1:2])

      }
      name <- vi_name[i]
      betas[[name]] <- a
      a<-rbind(colnames(a),a)

      #Anova Table:
      #b<-as.matrix(round(anova(tmp),3))
      b<-as.matrix(round(Anova(tmp),3))  #Need 'car' package
      b<-rbind(colnames(b),b)
      #a<-round(coefficients(summary(tmp))[2,],3)
      re<-rbind(re,
                c(colnames(vi)[i],rep("",3)),
                c("Coefficients:",rep("",3)),
                a,
                c("ANOVA table:",rep("",3)),
                b,
                "")
      colnames(re)<-rep("",4)

      #Result table:
      temp1<-a[3:(nlevels(group)+1),]
      if(CI==T){
        ci<-paste('(',round(confint(tmp)[,1],1),', ', round(confint(tmp)[,2],1),')',sep = "")
        civ<-ci[2:nlevels(group)]
        if(nlevels(group)==2){
          evi<-paste(round(as.numeric(temp1[1]),2),civ)
          tb<-c("0",evi,temp1[4])
        }else{
          evi<-paste(round(as.numeric(temp1[,1]),2),civ)
          tb<-c("0",evi,b[2,4])
        }
      }else{
        if(is.vector(temp1)==T){
          temp2<-t(rbind(c(0,"-"),
                         cbind(paste(round(as.numeric(temp1[1]),2),' ± ',round(as.numeric(temp1[2]),2),sep = ""),
                               temp1[4])))
        }else{
          temp2<-t(rbind(c(0,"-"),
                         cbind(paste(round(as.numeric(temp1[,1]),2),' ± ',round(as.numeric(temp1[,2]),2),sep = ""),
                               temp1[,4])))
        }
        if(nlevels(group)>2){
          tb<-c(t(temp2[1,]),b[2,4])
        }else{
          tb<-c(t(temp2[1,]),temp1[4])
        }
      }
      tb[2:nlevels(group)]<-paste(tb[2:nlevels(group)],ifelse(as.numeric(a[3:(nlevels(group)+1),4])>0.05,
                                                              "",
                                                              ifelse(as.numeric(a[3:(nlevels(group)+1),4])<=0.05 & as.numeric(a[3:(nlevels(group)+1),4])>0.01,
                                                                     "*",
                                                                     "**")))
      re1<-rbind(re1,tb)
    }


    re1[,ncol(re1)]<-ifelse(as.numeric(re1[,ncol(re1)])>=0.05,
                            re1[,ncol(re1)],
                            ifelse(as.numeric(re1[,ncol(re1)])<0.05 & as.numeric(re1[,ncol(re1)])>=0.01,
                                   paste(re1[,ncol(re1)],"*"),
                                   ifelse(as.numeric(re1[,ncol(re1)])<0.01 & as.numeric(re1[,ncol(re1)])>=0.001,
                                          paste(re1[,ncol(re1)],"**"), "<0.001 **")))


    rownames(re1)<-colnames(vi)
    colnames(re1)<-c(paste(levels(group), " vs. ",ref," (Mean Difference)",sep = ""),
                     "P value")
    re<-cbind(rownames(re),re)
    l<-list(re1,re)
    #Formating:
    #caption:

    if(CI==T){
      if (nlevels(group)>2) {
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ". P Value were obtained from F test. '*':  P (Difference comparing to the reference group) = 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }else{
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ".  '*':  P (Difference comparing to the reference group) < 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }
    }else{
      if (nlevels(group)>2) {
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ". P Value were obtained from F test. '*':  P (Difference comparing to the reference group) = 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }else{
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ".  '*':  P (Difference comparing to the reference group) < 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }
    }

    addWorksheet(wb,"Coefficients & ANOVA")
    writeData(wb,sheet = "Coefficients & ANOVA",re)
    saveWorkbook(wb,outfile,overwrite = T)
    openXL(outfile)
  }
  return(l)
}

#' Post Hoc Analysis Table
#'
#' Perform post hoc pairwise comparisons between treatment groups after ANOVA.
#'
#' @param vi A matrix or data frame with continuous variables of interest.
#' @param group A vector indicating group membership.
#' @param methods Post hoc method (e.g., "hsd", "bonf", "lsd", etc.).
#' @param based.on.aov Logical; whether to include only significant ANOVA results.
#' @param outfile Output Excel file path.
#'
#' @return A data frame of pairwise comparison results with p-values and CIs.
#' @export
PostHoc <- function(vi, group, methods = "hsd", based.on.aov = FALSE, outfile = 'Post Hoc Analysis.xlsx') {
  tb<-c()
  anova<-c()
  RE<-c()
  vi<-apply(apply(vi,2,as.character), 2, as.numeric)
  for (i in 1:ncol(vi)) {
    test<-cbind.data.frame(vi[,i], group)
    test[,1]<-as.character(as.numeric(test[,1]))
    test.aov<-aov(`vi[, i]`~group,data=test)
    anova<-c(anova,unlist(summary(test.aov))[9])
    re<-PostHocTest(test.aov, adj.methods=methods)$group
    a<-t(PostHocTest(test.aov, adj.methods=methods)$group)
    b<-paste(round(a[1,],2), ' (', round(a[2,],1), ', ', round(a[3,],1), ')', sep = '')
    adj.p<-a[4,]
    b[which(adj.p<0.05)]<-paste(b[which(adj.p<0.05)], '*')
    tb<-rbind(tb,b)

    rn<-c(colnames(vi)[i],paste("          ",rownames(re)))
    re<-rbind("",round(re,2))
    rownames(re)<-rn
    RE<-rbind(RE,re,"")
  }
  rownames(tb)<-colnames(vi)
  colnames(tb)<-paste(colnames(a), " (95% CI)", sep = "")
  colnames(RE)<-c("Mean Difference", "Lower 95% CI", "Higher 95% CI", "P Values")

  if(based.on.aov==T){
    tb<-tb[which(anova<0.05),]
  }

  ExcelTable(RE,
             title="Post Hoc Analysis",
             caption=paste( methods, "method was used to correct for multiple comparison"),
             outfile=outfile,
             openfile=T)
  return(RE)
}
