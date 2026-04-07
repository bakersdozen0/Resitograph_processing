# AUTOMATED RESISTOGRAPH PROCESSING SUITE
# Version: Streamlined Production Pipeline
# Description: Automates .rgp processing, generates visual reports, and exports to Excel.

# SECTION 1: SETUP & LIBRARIES ####
library(here)
library(tools)
library(dplyr)
library(ggplot2)
library(tidyr)
library(readxl)
library(gridExtra)
library(openxlsx)


# SECTION 2: ROBUST MATH HELPERS ####

safe_max <- function(x) {
  x <- na.omit(x)
  if(length(x) == 0) return(-Inf)
  return(max(x))
}

safe_min <- function(x) {
  x <- na.omit(x)
  if(length(x) == 0) return(Inf)
  return(min(x))
}

safe_var <- function(x) {
  x <- na.omit(x)
  if(length(x) < 2) return(0)
  return(var(x))
}


# SECTION 3: CORE PROCESSING ALGORITHMS #####

tan.linear.correct <- function(meanresistance, feedspeed, RPM){
  if(as.numeric(feedspeed) == 150 && as.numeric(RPM) == 2500){
    tana <- 0.03168201*exp(0.08160494*meanresistance)
  } else if (as.numeric(feedspeed) == 200 && as.numeric(RPM) == 2500){
    tana <- (0.03168201-0.01494255)*exp((0.08881194+0.01228310)*meanresistance)
  } else {
    tana <- 0
  }
  return(tana)
}

standardisetrace <- function(data1, feedspeed, RPM){
  FRcoef <- (as.numeric(feedspeed)*10* 0.00032624) + 0.34495 
  RPMcoef<- (-0.768 * log(2500/as.numeric(RPM)) + 1) 
  Stcoef <- (FRcoef / RPMcoef)
  data2 <- data1/Stcoef
  return(data2)
}

finding.cork <- function(combined_value,nmeasuredpoint,length_value,
                         tan.linear.correct,speedfeed, speeddrill){
  i <- 0
  if(nrow(combined_value) < 100) return(c(0, length_value)) 
  
  drillsubset <- combined_value[(length(combined_value$drill)-i):
                                  (length(combined_value$drill)-i-49), 1]
  
  if (safe_var(drillsubset)>1/8){
    while (safe_var(drillsubset)>1/8 && i <100) {
      i <- i+10
      drillsubset <- combined_value[(length(combined_value$drill)-i):
                                      (length(combined_value$drill)-i-49), 1]
    }
  }
  if (i == 100 && safe_var(drillsubset)>1/8){ i <- 0 }
  
  while (safe_var(drillsubset)<1/8 && i < (length(combined_value$drill)-50)) {
    i <- i+1
    drillsubset <- combined_value[(length(combined_value$drill)-i):
                                    (length(combined_value$drill)-i-49), 1]
  }
  
  if (i >0){
    meanplateau <- mean(combined_value[c((length(combined_value$drill)-i):
                                           length(combined_value$drill)),1], na.rm=TRUE)
    continualbark <- c()
    p <-1
    max_p <- length(combined_value$drill) - i - 1
    
    while (p < max_p && combined_value[(length(combined_value$drill)-i-p),1]< 2*meanplateau){
      idx <- length(combined_value$drill)-i-p+1
      mid_idx <- floor(length(combined_value$drill)/2)
      
      if (combined_value[idx,1] < meanplateau+1 &&
          combined_value[idx,3] > combined_value[mid_idx,3]){
        continualbark <- c(continualbark, combined_value[idx,3])
      }
      p <- p+1
    }
    
    continualbark2 <- c()
    x<- 0
    if (length(continualbark)>1){
      for (p in 1:(length(continualbark)-1)){
        if (abs(continualbark[p+1]-continualbark[p])<0.031){
          x <- x+1
        }else{
          if (x>9){
            continualbark2 <- c(continualbark2, continualbark[p])
            x <- 0
          } else{ x <- 0 }
        }
      }
    }
    if (x>9){ continualbark2 <- c(continualbark2, continualbark[length(continualbark)]) }
  } else {
    continualbark2 <- c()
  }
  
  if (i>0){
    if (length(continualbark2)>0){
      outcorkvalue <- min(continualbark2)
    }else{
      outcorkvalue <- (length(combined_value$drill)-i-50)/nmeasuredpoint*length_value
    }
  } else {
    discpoints <-seq(1*nmeasuredpoint/length_value, nmeasuredpoint, 1)
    discpoints <- discpoints[discpoints <= nrow(combined_value)]
    resistancerawcork <- mean(combined_value[discpoints,1], na.rm=TRUE)
    tana<- tan.linear.correct(resistancerawcork,speedfeed,speeddrill)
    p <- combined_value[nmeasuredpoint,1]-(tana*combined_value[nmeasuredpoint,3])
    
    if (!is.na(p) && abs((combined_value[1,1]-p))<4){
      outcorkvalue <- (nmeasuredpoint-1)/nmeasuredpoint*length_value
    }else {
      outcorkvalue <- length_value
    }
  }
  
  i <- 0
  drillsubset <- combined_value[(1+i):(10+i), 1]
  while (safe_var(drillsubset)<1/30 && i < 100) {
    i <- i+1
    drillsubset <- combined_value[(1+i):(10+i), 1]
  }
  entcorkvalue <- (1+i)/nmeasuredpoint*length_value
  return(c(entcorkvalue, outcorkvalue))
}

finding.peaks <- function(slopeconsecutivetotal,entervalue, exitvalue,
                          combined_value, length_value, nmeasuredpoint, meanvalue){
  
  peak_ind <- which(diff(sign(slopeconsecutivetotal)) != 0) + 1
  if(length(peak_ind) == 0) return(list(c(), c())) 
  
  peakcoordinate <- numeric(length(peak_ind) + 2)
  peakcoordinate[2:(length(peak_ind)+1)] <- combined_value[peak_ind, 3]
  peakcoordinate[1]<- combined_value[entervalue,3]
  peakcoordinate[length(peakcoordinate)] <-  combined_value[exitvalue,3]
  
  peakcoordinate <- peakcoordinate*nmeasuredpoint/length_value
  
  start_idx <- max(1, floor(entervalue+2*nmeasuredpoint/length_value))
  end_idx <- min(nrow(combined_value), floor(exitvalue-2*nmeasuredpoint/length_value))
  
  if(start_idx < end_idx){
    thresholdvalue <- safe_var(combined_value[start_idx:end_idx, 4])
  } else {
    thresholdvalue <- 0
  }
  
  if (thresholdvalue/10 < meanvalue/6){ 
    thresholdvalue <- 10*meanvalue/6 
  }
  
  peakcoordinate <- round(peakcoordinate, digits = 0)
  peakcoordinate <- peakcoordinate[peakcoordinate > 0 & peakcoordinate <= nrow(combined_value)]
  
  peakcoordinate2 <- c()
  signvalue <- -1
  previouspeak <- peakcoordinate[1]
  
  if(length(peakcoordinate) > 1){
    for(i in 1:(length(peakcoordinate)-1)){
      val_next <- combined_value[peakcoordinate[i+1],4]
      val_prev <- combined_value[previouspeak,4]
      if(is.na(val_next) || is.na(val_prev)) next
      
      if (abs(val_next - val_prev) < (thresholdvalue/10)){   
        if (signvalue <0){
          if(val_next < val_prev){ previouspeak <- peakcoordinate[i+1] }
        }else {
          if(val_next > val_prev){ previouspeak <- peakcoordinate[i+1] }
        }
      } else {
        peakcoordinate2 <- c(peakcoordinate2, previouspeak)
        signvalue <- sign(val_next - val_prev)
        previouspeak <- peakcoordinate[i+1]
      }
    }
  }
  peakcoordinate2 <- c(peakcoordinate2, previouspeak)
  
  if(length(peakcoordinate2) > 2){
    deletevalue <- c()
    for (i in seq_along(peakcoordinate2)[-c(length(peakcoordinate2), length(peakcoordinate2)-1)]){
      s1 <- sign(combined_value[peakcoordinate2[i+1],4]-combined_value[peakcoordinate2[i],4])
      s2 <- sign(combined_value[peakcoordinate2[i+2],4]-combined_value[peakcoordinate2[i+1],4])
      if (!is.na(s1) && !is.na(s2) && s1==s2){
        deletevalue <- c(deletevalue,i+1)
      }
    }
    if (length(deletevalue)>0){ peakcoordinate2 <- peakcoordinate2[-deletevalue] }
  }
  return(list(peakcoordinate,peakcoordinate2))
}

finding.entrybark <- function(peakcoordinate2, combined_value, length_value, 
                              nmeasuredpoint,meanvalue, entcorkvalue, 
                              slopeconsecutivetotal2, peakcoordinate){
  peakenter <- c()
  peakexit <- c()
  
  if(length(peakcoordinate2) > 1){
    for (i in 1:(length(peakcoordinate2)-1)){
      intvalue <- combined_value[peakcoordinate2[i+1],4]-combined_value[peakcoordinate2[i],4]
      if (!is.na(intvalue) && intvalue <0){ peakexit <- c(peakexit, peakcoordinate2[i+1], intvalue)
      } else if (!is.na(intvalue)) { peakenter <- c(peakenter, peakcoordinate2[i], intvalue) }
    }
  }
  
  peakenter2 <- c()
  if(length(peakenter) > 0){
    ascending <-seq(1, length(peakenter),2)
    check_val <- (entcorkvalue+2.5)*nmeasuredpoint/length_value
    if(any(peakenter[ascending] < check_val)){
      limit_idx <- max(which(peakenter[ascending] < check_val))
      peakenter2 <- peakenter[1:(limit_idx*2)]
    }
  }
  
  deletevalue <- c()
  if (length(peakenter2)>=4){
    for (i in 1:((length(peakenter2)/2)-1)){
      idx1 <- (2*i)-1
      idx2 <- 2+(2*i)-1
      if (abs(combined_value[peakenter2[idx1],4]-combined_value[peakenter2[idx2],4]) < (meanvalue/2.5)){
        deletevalue<- c(deletevalue, c(((2*i)-1),(2*i)))
      }else{ break }
    }
  }
  if (length(deletevalue)>0){ peakenter2 <- peakenter2[-deletevalue] }
  
  barkenterpeak <- numeric()
  if(length(peakenter2) > 0 && length(peakenter) > 0){
    med_peak <- median(peakenter[seq(1, length(peakenter), 2)], na.rm=TRUE) 
    for (i in 1:(length(peakenter2)/2)){
      if (peakenter2[2*i] > med_peak){
        barkenterpeak <- peakenter2[(2*i)-1]
        break
      }
    }
    if (length(barkenterpeak)==0){
      if(length(peakenter2) >= 2){
        seq_peaks <- peakenter2[seq(2,length(peakenter2),2)]
        i <- which(peakenter2==max(seq_peaks, na.rm=TRUE))[1]
        barkenterpeak <- peakenter2[i-1]
      }
    }
  }
  
  if(length(barkenterpeak) == 0) return(list(1, peakexit)) 
  
  peakenter2 <- c(barkenterpeak, peakcoordinate2[which(peakcoordinate2==barkenterpeak)[1]+1])
  rng <- which(peakcoordinate == min(peakenter2))[1] : which(peakcoordinate == max(peakenter2))[1]
  peakenter2 <- peakcoordinate[rng]
  
  increment <- diff(combined_value[peakenter2,4])
  if(length(increment) > 0){
    idx_max <- which(increment == max(increment, na.rm=TRUE))[1]
    peakenter2 <- peakenter2[c(idx_max, idx_max+1)]
    
    barkenter <- max(slopeconsecutivetotal2[c(peakenter2[1]:peakenter2[2])], na.rm=TRUE)
    barkenter3 <- combined_value[(barkenter==slopeconsecutivetotal2),3]
    valid_mask <- (barkenter3*nmeasuredpoint/length_value > min(peakenter2)) &
      (barkenter3*nmeasuredpoint/length_value < max(peakenter2))
    barkenter <- barkenter3[valid_mask]
    if(length(barkenter) > 1) barkenter <- min(barkenter)
    if(length(barkenter) == 0) barkenter <- peakenter2[1]
  } else {
    barkenter <- peakenter2[1]
  }
  
  return(list(barkenter, peakexit))
}

finding.exitbark <- function(peakcoordinate2, combined_value, length_value, nmeasuredpoint,meanvalue,
                             outcorkvalue, slopeconsecutivetotal2, peakcoordinate, peakexit){
  
  peakexit2 <- c()
  if(length(peakexit) > 0){
    ascending <-seq(1, length(peakexit),2)
    check_val <- (outcorkvalue-2.5)*nmeasuredpoint/length_value
    if(any(peakexit[ascending] > check_val)){
      start_idx <- (min(which(peakexit[ascending] > check_val))*2-1)
      peakexit2 <- peakexit[start_idx:length(peakexit)]
    }
    if (length(peakexit2)==0 && length(peakexit) >= 2){
      peakexit2 <- c(peakexit[length(peakexit)-1], peakexit[length(peakexit)])
    }
  }
  
  barkexitpeak <- numeric()
  if(length(peakexit2) > 0 && length(peakexit) > 0){
    med_peak <- median(peakexit[-seq(1, length(peakexit),2)], na.rm=TRUE)
    for (i in 1:(length(peakexit2)/2)){
      idx <- (length(peakexit2)+2)-(2*i)
      if (peakexit2[idx] < med_peak){
        barkexitpeak <- peakexit2[idx-1]
        break
      }
    }
  }
  
  if (length(barkexitpeak)==0 && length(peakexit2) >= 2){
    seq_vals <- peakexit2[seq(2,length(peakexit2),2)]
    i <- which(peakexit2==min(seq_vals, na.rm=TRUE))[1]
    barkexitpeak <- peakexit2[i-1]
  }
  
  if(length(barkexitpeak) == 0) return(length_value)
  
  idx_peak <- which(peakcoordinate2==barkexitpeak)[1]
  if(is.na(idx_peak) || idx_peak == 1) return(barkexitpeak)
  
  peakexit2 <- c(peakcoordinate2[idx_peak-1], barkexitpeak)
  rng <- which(peakcoordinate == min(peakexit2))[1] : which(peakcoordinate == max(peakexit2))[1]
  peakexit2 <- peakcoordinate[rng]
  
  increment <- diff(combined_value[peakexit2,4])
  if(length(increment) > 0){
    idx_min <- which(increment == min(increment, na.rm=TRUE))[1]
    peakexit2 <- peakexit2[c(idx_min, idx_min+1)]
    
    barkexit <- max(abs(slopeconsecutivetotal2[c(peakexit2[1]:(peakexit2[2]-1))]), na.rm=TRUE)
    barkexit3 <- combined_value[(barkexit == slopeconsecutivetotal2),3]
    valid_mask <- (barkexit3*nmeasuredpoint/length_value > min(peakexit2)) & 
      (barkexit3*nmeasuredpoint/length_value < max(peakexit2))
    barkexit <- barkexit3[valid_mask]
    if(length(barkexit) > 1) barkexit <- max(barkexit)
    if(length(barkexit) == 0) barkexit <- barkexitpeak
  } else {
    barkexit <- barkexitpeak
  }
  return(barkexit)
}

finding.cracks <- function(peakcoordinate2, entervalue2, exitvalue2, combined_value, meanvalue){
  potentialcracks <- peakcoordinate2[(peakcoordinate2 > entervalue2 + 50) & (peakcoordinate2 < exitvalue2 - 50) & (combined_value[peakcoordinate2, 4] < meanvalue / 8)]
  ifelse(length(potentialcracks) > 0, cracklocation <-c(combined_value[potentialcracks, 3],200,300), cracklocation <- c(200,300))
  cracks <- length(cracklocation)-2
  return(list(cracks, cracklocation))
}

finding.pith <-function(DBHcore,outcorkvalue, length_value, nmeasuredpoint,pitlocation,peakcoordinate2, combined_value){
  
  # Return default if missing DBH
  if(is.na(DBHcore)) return(pitlocation)
  
  if(DBHcore < 25) { dividevalue <- 3 } else { dividevalue <- 5 }
  
  if(outcorkvalue!=length_value){
    pitarea <-round(DBHcore/dividevalue, digits = 0)*nmeasuredpoint/length_value
    pitarea <- round(c(pitlocation-pitarea, pitlocation+pitarea), digits = 0)
  } else{
    pitarea <-trunc(DBHcore/dividevalue)*nmeasuredpoint/length_value
    pitarea <- round(c(pitlocation-pitarea, pitlocation+1.5*pitarea), digits = 0)
  }
  
  peakcoordinatepeak <-peakcoordinate2[seq(2,length(peakcoordinate2),2)]
  peakcoordinatevalley <-peakcoordinate2[seq(1,length(peakcoordinate2),2)]
  
  if(length(peakcoordinatepeak) > 0) {
    peakcoordinatepeak <- peakcoordinatepeak[(peakcoordinatepeak>min(pitarea))& (peakcoordinatepeak<max(pitarea))]
  }
  if(length(peakcoordinatevalley) > 0) {
    peakcoordinatevalley <- peakcoordinatevalley[(peakcoordinatevalley>min(pitarea))& (peakcoordinatevalley<max(pitarea))]
  }
  
  potentialpit <-c()
  if(length(peakcoordinatepeak) > 1 && length(peakcoordinatevalley) > 1){
    widepeak <- diff(peakcoordinatepeak)
    widepeak2 <- diff(peakcoordinatevalley)
    
    max_w1 <- safe_max(widepeak)
    max_w2 <- safe_max(widepeak2)
    
    if (max_w2>180 && max_w1>160){
      if(max_w1 > max_w2){
        if (safe_min(peakcoordinatepeak) < safe_min(peakcoordinatevalley)){ 
          potentialpit <- c(potentialpit, peakcoordinatevalley[which.max(widepeak)])
        }else{ 
          potentialpit <- c(potentialpit, peakcoordinatevalley[which.max(widepeak)+1]) 
        }
      }else {
        if (safe_min(peakcoordinatepeak) < safe_min(peakcoordinatevalley)){ 
          potentialpit <- c(potentialpit, peakcoordinatepeak[which.max(widepeak2)+1])
        }else{ 
          potentialpit <- c(potentialpit, peakcoordinatepeak[which.max(widepeak2)]) 
        }
      }
    } 
  }
  
  if(length(potentialpit)<1){
    potentialpit <- pitlocation
  } else {
    potentialpit <- mean(potentialpit, na.rm=TRUE)
  }
  return(potentialpit)
}

finding.ring <- function(peakcoordinate2, potentialpit, entervalue2, exitvalue2, slopeconsecutivetotal){
  peaks <-peakcoordinate2[seq(2,length(peakcoordinate2),2)] 
  ringcoordinateentry <- c()
  ringcoordinateexit <- c()
  
  if(length(peaks) > 0){
    for (i in 1:length(peaks)){
      if(peaks[i]>entervalue2 && peaks[i]<potentialpit){
        ringcoordinateentry <-  c(ringcoordinateentry, peakcoordinate2[which(peaks[i]==peakcoordinate2)-1], peaks[i])
      } else if (peaks[i]>potentialpit && peaks[i]<exitvalue2){
        ringcoordinateexit <- c(ringcoordinateexit, peaks[i], peakcoordinate2[which(peaks[i]==peakcoordinate2)+1])
      }
    }
  }
  
  if(is.null(ringcoordinateexit) || length(ringcoordinateexit) <= 2){
    if(length(peaks) >= 2){
      ringcoordinateexit <- c(peaks[length(peaks)-1], peakcoordinate2[length(peakcoordinate2)-1],
                              peaks[length(peaks)], peakcoordinate2[length(peakcoordinate2)])
    }
  }
  
  ringcoordinateentry2 <- c()
  ringcoordinateexit2 <- c()
  
  if(length(ringcoordinateentry) > 1){
    for(i in 1:(length(ringcoordinateentry)/2)){
      intvalue <-slopeconsecutivetotal[ringcoordinateentry[((2*i)-1)]: ringcoordinateentry[2*i]]
      if(length(intvalue) > 0)
        ringcoordinateentry2 <- c(ringcoordinateentry2, which(intvalue==max(intvalue, na.rm=TRUE))+ ringcoordinateentry[((2*i)-1)])
    }
  }
  if(length(ringcoordinateexit) > 1){
    for(i in 1:((length(ringcoordinateexit)/2)-1)){
      intvalue <-slopeconsecutivetotal[ringcoordinateexit[((2*i)-1)]: ringcoordinateexit[2*i]]
      if(length(intvalue) > 0)
        ringcoordinateexit2 <- c(ringcoordinateexit2, which(intvalue==min(intvalue, na.rm=TRUE))+ ringcoordinateexit[((2*i)-1)])
    }
  }
  if(length(ringcoordinateentry2) > 0) ringcoordinateentry2 <- ringcoordinateentry2[-1]
  ringcoordinate <- c(ringcoordinateentry2, ringcoordinateexit2)
  return(ringcoordinate)
}


# SECTION 4: MAIN FILE PARSING WRAPPER #### 

process_single_rgp <- function(filepath, filename, standard="no", return_plot = FALSE){  
  # 1. Read File
  resistograph_rgp <- scan(filepath, what = "", quiet = TRUE)
  if(length(resistograph_rgp) < 10) stop("Empty or corrupted file")
  
  # 2. Extract Data
  drill_values <- resistograph_rgp[grep("drill", resistograph_rgp)+2]
  feed_values <- resistograph_rgp[grep("drill", resistograph_rgp)+5]
  length_value <- resistograph_rgp[grep("depthMsmt", resistograph_rgp)+2]
  
  drill_values <- substring(drill_values, 2, nchar(drill_values)-2)
  drill_values2 <- as.numeric(unlist(strsplit(drill_values,",")))
  
  feed_values <- substring(feed_values, 2, nchar(feed_values)-2)
  feed_values2 <- as.numeric(unlist(strsplit(feed_values,",")))
  
  length_value <- as.numeric(substring(length_value, 1, nchar(length_value)-1))
  length_intermediate <- seq(length_value/length(feed_values2), length_value, length_value/length(feed_values2))
  
  combined_value <- data.frame(drill_values2, feed_values2, length_intermediate)
  colnames(combined_value) <- c("drill","feed", "length")
  nmeasuredpoint <- length(combined_value$drill)
  
  # 3. Extract Metadata
  tree_ID <- resistograph_rgp[grep("idNumber", resistograph_rgp)+2]
  
  date_time_raw <- c(resistograph_rgp[grep("dateDay", resistograph_rgp)+2],
                     resistograph_rgp[grep("dateMonth", resistograph_rgp)+2],
                     resistograph_rgp[grep("dateYear", resistograph_rgp)+2],
                     resistograph_rgp[grep("timeHour", resistograph_rgp)+2],
                     resistograph_rgp[grep("timeMinute", resistograph_rgp)+2],
                     resistograph_rgp[grep("timeSecond", resistograph_rgp)+2])
  date_time_raw <- gsub(',','',date_time_raw)
  date_time_raw <- ifelse(nchar(date_time_raw) == 1, paste0("0", date_time_raw), date_time_raw)
  date_time <- paste(paste(date_time_raw[1:3], collapse="/"), paste(date_time_raw[4:6], collapse=":"))
  
  speedfeed<- gsub(',','',resistograph_rgp[grep("speedFeed", resistograph_rgp)+2])
  speeddrill<- gsub(',','', resistograph_rgp[grep("speedDrill", resistograph_rgp)+2])
  offsetfeed <- gsub(',','', resistograph_rgp[grep("offsetFeed", resistograph_rgp)+2])
  offsetdrill <- gsub(',','', resistograph_rgp[grep("offsetDrill", resistograph_rgp)+2])
  
  # 4. Standardisation
  if(standard=="yes"){
    combined_value$drill <- standardisetrace(combined_value$drill, speedfeed, speeddrill)
    combined_value$feed <- standardisetrace(combined_value$feed, speedfeed, speeddrill)
  }
  
  # 5. Core Logic
  entcorkvalue_all <- finding.cork(combined_value,nmeasuredpoint,length_value,
                                   tan.linear.correct,speedfeed,speeddrill)
  outcorkvalue <- entcorkvalue_all[2]
  entcorkvalue <- entcorkvalue_all[1]
  
  Drill_Status <- "Complete"
  
  if (as.numeric(outcorkvalue) > as.numeric(length_value)) { outcorkvalue <- length_value }
  
  # Initialize numeric placeholders
  DBHdiscs <- NA_real_
  DBHcore <- NA_real_
  barkthicknessout <- NA_real_
  
  exitvalue <- round(outcorkvalue*nmeasuredpoint/length_value, digits = 0)
  entervalue <- round(entcorkvalue*nmeasuredpoint/length_value, digits = 0)
  
  discpoints <-seq(entcorkvalue*nmeasuredpoint/length_value, outcorkvalue*nmeasuredpoint/length_value, 1)
  discpoints <- discpoints[discpoints > 0 & discpoints <= nrow(combined_value)]
  resistancerawcork <- mean(combined_value[discpoints,1], na.rm=TRUE)
  
  combined_value$drillcorrected <- NA
  if (combined_value[1,1]<= combined_value[exitvalue,1] && outcorkvalue!=length_value){
    tana <- (combined_value[exitvalue,1]-combined_value[1,1])/outcorkvalue
    combined_value$drillcorrected[1:exitvalue] <- combined_value$drill[1:exitvalue] - (tana*combined_value$length[1:exitvalue])
  } else {
    tana <- tan.linear.correct(resistancerawcork,speedfeed,speeddrill)
    combined_value$drillcorrected[1:exitvalue] <- combined_value$drill[1:exitvalue] - (tana*combined_value$length[1:exitvalue])
  }
  
  slopeconsecutivetotal <- diff(combined_value[(1:exitvalue),4])/diff(combined_value[(1:exitvalue),3])
  slopeconsecutivetotal2 <- abs(slopeconsecutivetotal)
  
  # Safe Mean
  meanvalue <- mean(combined_value[(entervalue+2*nmeasuredpoint/length_value):
                                     (exitvalue-2*nmeasuredpoint/length_value), 4], na.rm=TRUE)
  if(is.nan(meanvalue) || is.na(meanvalue)) meanvalue <- 0
  
  peak_res <- finding.peaks(slopeconsecutivetotal, entervalue, exitvalue, 
                            combined_value, length_value, nmeasuredpoint,meanvalue)
  peakcoordinate2 <- unlist(peak_res[2])
  peakcoordinate <- unlist(peak_res[1])
  
  bark_res <- finding.entrybark(peakcoordinate2, combined_value, length_value, nmeasuredpoint,
                                meanvalue, entcorkvalue, slopeconsecutivetotal2, peakcoordinate)
  barkenter <- unlist(bark_res[1])
  peakexit_temp <- unlist(bark_res[2])
  
  barkexit <- NA
  if(outcorkvalue!=length_value){
    barkexit <- finding.exitbark(peakcoordinate2, combined_value, length_value, nmeasuredpoint,
                                 meanvalue, outcorkvalue, slopeconsecutivetotal2, peakcoordinate, peakexit_temp)
  }
  
  entervalue2 <-barkenter*nmeasuredpoint/length_value
  
  if(outcorkvalue!=length_value){
    exitvalue2 <- barkexit*nmeasuredpoint/length_value
    DBHcore <- barkexit-barkenter
    DBHdiscs <- outcorkvalue - entcorkvalue
    barkthicknessout <- outcorkvalue - barkexit
  } else {
    exitvalue2 <- exitvalue
    Drill_Status <- "Incomplete"
  }
  
  barkthicknessin <- barkenter - entcorkvalue
  pitlocation <- entervalue2+(exitvalue2-entervalue2)/2
  
  cracks_res <- finding.cracks(peakcoordinate2, entervalue2, exitvalue2, combined_value, meanvalue)
  cracks <- unlist(cracks_res[1])
  
  potentialpit <- finding.pith(as.numeric(DBHcore),outcorkvalue, length_value, nmeasuredpoint,
                               pitlocation, peakcoordinate2, combined_value)
  
  entryradius <- (potentialpit - entervalue2)*length_value/nmeasuredpoint
  
  if(!is.na(potentialpit)) {
    ringcoordinate <- finding.ring(peakcoordinate2, potentialpit, entervalue2, exitvalue2, slopeconsecutivetotal)
    ringnumber <- length(which(ringcoordinate < (potentialpit-1*nmeasuredpoint/length_value)))+1
  } else {
    ringnumber <- NA
  }
  
  resistanceRaw <- mean(combined_value[(entervalue2:exitvalue2),1], na.rm=TRUE)
  resistancecorrected <- mean(combined_value[(entervalue2:exitvalue2),4], na.rm=TRUE)
  resistancecorrectedcork <- mean(combined_value[discpoints,4], na.rm=TRUE)
  resistance3 <- mean(combined_value[(entervalue2:(entervalue2+3*nmeasuredpoint/length_value)),4], na.rm=TRUE)
  resistance6 <- mean(combined_value[(entervalue2:(entervalue2+6*nmeasuredpoint/length_value)),4], na.rm=TRUE)
  resistancepit <- mean(combined_value[(entervalue2:potentialpit),4], na.rm=TRUE)
  
  # --- NEW PLOTTING BLOCK ---
  plot_obj <- NULL
  if (return_plot) {
    # Clean up dummy crack coordinates (200, 300) set by finding.cracks
    crack_locs <- cracks_res[[2]]
    crack_locs <- crack_locs[crack_locs < length_value] 
    
    # DOWNSAMPLE: Keep only every 5th point to drastically speed up PDF rendering
    plot_data <- combined_value[seq(1, nrow(combined_value), by = 5), ]
    
    # Extract the parent directory name
    parent_folder <- basename(dirname(filepath))
    
    # Base ggplot mapping the trace
    p <- ggplot(plot_data, aes(x = length)) +
      geom_line(aes(y = drill), color = "ivory3", alpha = 0.8) +
      geom_line(aes(y = drillcorrected), color = "orange") +
      
      # Cork limits
      geom_vline(xintercept = entcorkvalue, color = "red", linetype = "dashed") +
      geom_vline(xintercept = outcorkvalue, color = "red", linetype = "dashed") +
      
      # Bark entry
      geom_vline(xintercept = barkenter, color = "darkred", linewidth = 1) +
      
      theme_minimal() +
      theme(panel.background = element_rect(fill = "white", color = NA),
            plot.background = element_rect(fill = "white", color = NA)) +
      labs(
        # UPDATE: Now prints Folder Name and File Name as the main title
        title = paste("Folder:", parent_folder, "| File:", filename),
        subtitle = paste("Feed:", speedfeed, "cm/min | Drill:", speeddrill, "RPM | Status:", Drill_Status),
        x = "Drilling depth (cm)", y = "Resistance amplitude (%)"
      )
    
    # Conditional plot layers (Bark exit, Pith, Rings, Cracks)
    if (!is.na(barkexit)) {
      p <- p + geom_vline(xintercept = barkexit, color = "darkred", linewidth = 1)
    }
    if (!is.na(potentialpit)) {
      p <- p + geom_vline(xintercept = potentialpit * length_value / nmeasuredpoint, color = "darkgreen", linetype = "dashed", linewidth = 1)
    }
    if (!is.na(potentialpit) && length(ringcoordinate) > 0) {
      p <- p + geom_vline(xintercept = ringcoordinate * length_value / nmeasuredpoint, color = "tan", linetype = "dashed")
    }
    if (length(crack_locs) > 0) {
      p <- p + geom_vline(xintercept = crack_locs, color = "slateblue4", linetype = "dotted", linewidth = 1)
    }
    
    plot_obj <- p
  }
  
  # --- UPDATED RETURN STATEMENT ---
  res_df <- data.frame(
    File_name = filename,
    Tree_ID = tree_ID,
    Sample_date = date_time,
    Speed_cm_min = as.numeric(speedfeed),
    Drill_speed_RPM = as.numeric(speeddrill),
    offsetfeed = offsetfeed,
    offsetdrill = offsetdrill,
    DBH_disc = as.numeric(DBHdiscs),
    DBHcore = as.numeric(DBHcore),
    Bark_thickness_entry = as.numeric(barkthicknessin),
    Bark_thickness_exit = as.numeric(barkthicknessout),
    entry_radius_cm = as.numeric(entryradius),
    Detrended_resistance_stripped = as.numeric(resistancecorrected),
    Detrended_resistance_bark = as.numeric(resistancecorrectedcork),
    Raw_resistance_stripped = as.numeric(resistanceRaw),
    Raw_resistance_bark = as.numeric(resistancerawcork),
    Cracks = cracks,
    Tangente = tana,
    Resistance_3cm = as.numeric(resistance3),
    Resistance_6cm = as.numeric(resistance6),
    Resistance_pith = as.numeric(resistancepit),
    Number_rings = ringnumber,
    Resistance_age0_0 = NA,
    Standardised = standard,
    Drill_Status = Drill_Status
  )
  
  if (return_plot) {
    return(list(metrics = res_df, plot = plot_obj))
  } else {
    return(res_df)
  }
}

# SECTION 5: MASTER BATCH PIPELINE ####

batch_process_directory <- function(dir_path) {
  files <- list.files(dir_path, pattern = "\\.rgp$", full.names = TRUE, recursive = TRUE)
  if(length(files) == 0) stop(paste("No .rgp files found in:", dir_path))
  
  print(paste("Processing", length(files), "files found in directory structure..."))
  
  results_list <- list()
  
  # --- NEW: Initialize the failure tracker ---
  failed_files_report <- data.frame(
    Folder = character(),
    File_Name = character(),
    Reason = character(),
    stringsAsFactors = FALSE
  )
  
  for(f in files){
    fname <- basename(f)
    parent_folder <- basename(dirname(f))
    
    # Check 1: Is the file completely empty? (0 bytes)
    if (file.info(f)$size == 0) {
      failed_files_report <- rbind(failed_files_report, data.frame(
        Folder = parent_folder,
        File_Name = fname,
        Reason = "Empty File (0 bytes)"
      ))
      next # Skip to the next file
    }
    
    # Check 2: Try to process the file and catch internal corruption errors
    res <- tryCatch({
      
      process_single_rgp(f, fname, standard = "no")
      
    }, error = function(e) {
      # Return the error message as a string instead of crashing
      return(paste("ERROR:", e$message))
    })
    
    # If the tryCatch returned our error string, log it and skip
    if (is.character(res) && grepl("^ERROR:", res)) {
      failed_files_report <- rbind(failed_files_report, data.frame(
        Folder = parent_folder,
        File_Name = fname,
        Reason = res
      ))
      next
    }
    
    # If successful, add to our final list
    if(!is.null(res)){
      res$Source_Sheet <- parent_folder
      results_list[[paste(parent_folder, fname, sep="_")]] <- res
    }
  }
  
  # --- NEW: Print Audit Report ---
  message("\n=======================================================")
  message("             RGP PROCESSING AUDIT REPORT               ")
  message("=======================================================")
  message(paste("Total files found in directories:", length(files)))
  message(paste("Successfully processed:          ", length(results_list)))
  message(paste("Failed/Skipped files:            ", nrow(failed_files_report)))
  
  if (nrow(failed_files_report) > 0) {
    message("\n[!] DETAILED FAILURE LIST:")
    print(failed_files_report)
    
    # Optional: Automatically save this report to a CSV so you don't lose it
    # write.csv(failed_files_report, file.path(dirname(dir_path), "Failed_RGP_Report.csv"), row.names = FALSE)
  } else {
    message("\nAll files processed successfully! No empty or corrupt files found.")
  }
  message("=======================================================\n")
  
  return(bind_rows(results_list))
}

run_production_multisheet <- function(root_dir, output_excel_path, ignore_pattern = NULL) {
  print(paste("Scanning root directory:", basename(root_dir)))
  all_data <- batch_process_directory(root_dir)
  if(nrow(all_data) == 0) stop("No data found in directory.")
  
  if(!is.null(ignore_pattern)) {
    all_data <- all_data %>% filter(!grepl(ignore_pattern, Source_Sheet, ignore.case = TRUE))
  }
  
  all_data <- all_data %>%
    mutate(across(c(DBHcore, DBH_disc, Detrended_resistance_stripped, Resistance_3cm), ~suppressWarnings(as.numeric(.))))
  
  sheet_list <- split(all_data, all_data$Source_Sheet)
  print(paste("Writing", length(sheet_list), "sheets to", output_excel_path, "..."))
  write.xlsx(sheet_list, file = output_excel_path)
  print("Excel Export Done!")
}

generate_resi_pdf_report <- function(dir_path, target_subdir = NULL, recursive_search = TRUE, output_pdf_path, ignore_pattern = "Archive") {
  
  # 1. Path construction logic
  search_path <- dir_path
  if (!is.null(target_subdir) && target_subdir != "") {
    search_path <- file.path(dir_path, target_subdir)
  }
  
  if (!dir.exists(search_path)) {
    stop(paste("The specified directory does not exist:", search_path))
  }
  
  # --- NEW: SMART OUTPUT NAMING LOGIC ---
  # If processing a specific folder, append its name to the PDF file to prevent overwriting
  final_pdf_path <- output_pdf_path
  if (!is.null(target_subdir) && target_subdir != "") {
    
    # Break apart the original file path
    out_dir <- dirname(output_pdf_path)
    base_name <- tools::file_path_sans_ext(basename(output_pdf_path))
    ext <- tools::file_ext(output_pdf_path)
    
    # Clean up the folder name just in case it contains slashes
    safe_subdir_name <- gsub("[\\\\/]", "_", target_subdir)
    
    # Stitch it back together (e.g., "Experiment_Visual_Report_FolderA.pdf")
    final_pdf_path <- file.path(out_dir, paste0(base_name, "_", safe_subdir_name, ".", ext))
  }
  # --------------------------------------
  
  # list.files now respects the recursive toggle
  files <- list.files(search_path, pattern = "\\.rgp$", full.names = TRUE, recursive = recursive_search)
  
  if (!is.null(ignore_pattern) && ignore_pattern != "") {
    files <- files[!grepl(ignore_pattern, files, ignore.case = TRUE)]
  }
  
  if(length(files) == 0) stop(paste("No valid .rgp files found in:", search_path))
  
  print(paste("Generating PDF for", length(files), "scans... This may take a moment."))
  
  # Open PDF device using the new dynamic path
  pdf(final_pdf_path, width = 11, height = 8.5) 
  
  for(f in files){
    fname <- basename(f)
    tryCatch({
      res <- process_single_rgp(f, fname, standard = "no", return_plot = TRUE)
      
      if(!is.null(res$plot)) {
        print(res$plot) 
      }
    }, error = function(e){
      print(paste("SKIPPING plot for", fname, ":", e$message))
    })
  }
  
  dev.off() # Close and save the PDF
  print(paste("Done! PDF report saved successfully to:", final_pdf_path))
}
# SECTION 6: MAIN EXECUTION BLOCK ####

# Instructions: Update the variables below for your specific experiment and run.
# NB: all 3 steps of this section needs to be re-run after changing "experiment_name" to update directory paths 

# 1. Define Experiment Name and Root Directory
experiment_name <- "Llandovery 31 RESI" # e.g., "Moray 66" or "Kintyre 18"
root_folder     <- here(experiment_name) 

# 2. Define Output Paths
output_excel <- here("Processed data", paste0(experiment_name, "_Automated_Output.xlsx"))
output_pdf   <- here("Processed data", paste0(experiment_name, "_Visual_Report.pdf"))

# 3. Directories to Ignore (Set to NULL if none)
## NB: If multiple, DO NOT USE VECTOR LIST -- use the or operator "|" (e.g.: folder_to_skip <- "RESENT DEC 10 FILES|Problem files")
folder_to_skip <- "Blks 16-18"

# --- EXECUTE PHASE 1: Data Extraction ####
run_production_multisheet(
  root_dir = root_folder, 
  output_excel_path = output_excel, 
  ignore_pattern = folder_to_skip
)


# --- EXECUTE PHASE 2: Visual Report (PDF) ####
# OPTION A: Plot EVERYTHING (Default)
generate_resi_pdf_report(dir_path = root_folder, output_pdf_path = output_pdf, ignore_pattern = folder_to_skip)

# OPTION B: Plot only a specific batch to save time
generate_resi_pdf_report(
  dir_path = root_folder, 
  target_subdir = "Blk 16", # <--- Put the folder name here
  recursive_search = FALSE,                     # <--- Turn off deep diving
  output_pdf_path = output_pdf, 
  ignore_pattern = folder_to_skip
)

