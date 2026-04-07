
# AUTOMATED RESISTOGRAPH PROCESSING SUITE
# Version: benchmarking: 
# compares automated results to manually processed ones 

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


# SECTION 5: AUTOMATION HELPERS & TEST RUNNERS ####

# Helper: Recursively find and process all .rgp files AND capture folder name
batch_process_directory <- function(dir_path) {
  files <- list.files(dir_path, pattern = "\\.rgp$", full.names = TRUE, recursive = TRUE)
  if(length(files) == 0) stop(paste("No .rgp files found in:", dir_path))
  
  print(paste("Processing", length(files), "files found in directory structure..."))
  
  results_list <- list()
  for(f in files){
    fname <- basename(f)
    # Extract the immediate parent folder name (e.g., "Block 1")
    parent_folder <- basename(dirname(f))
    
    tryCatch({
      res <- process_single_rgp(f, fname, standard = "no")
      
      # Add the folder name to the result
      if(!is.null(res)){
        res$Source_Sheet <- parent_folder
        results_list[[paste(parent_folder, fname, sep="_")]] <- res
      }
    }, error = function(e){
      print(paste("SKIPPING", fname, ":", e$message))
    })
  }
  return(bind_rows(results_list))
}

generate_resi_pdf_report <- function(dir_path, output_pdf_path, ignore_pattern = "Archive") {
  files <- list.files(dir_path, pattern = "\\.rgp$", full.names = TRUE, recursive = TRUE)
  
  if (!is.null(ignore_pattern) && ignore_pattern != "") {
    files <- files[!grepl(ignore_pattern, files, ignore.case = TRUE)]
  }
  
  if(length(files) == 0) stop(paste("No valid .rgp files found in:", dir_path))
  
  print(paste("Generating PDF for", length(files), "scans... This may take a moment."))
  
  # Open PDF device (Landscape orientation for wide traces)
  pdf(output_pdf_path, width = 11, height = 8.5) 
  
  for(f in files){
    fname <- basename(f)
    tryCatch({
      # Process the file and request the plot object
      res <- process_single_rgp(f, fname, standard = "no", return_plot = TRUE)
      
      if(!is.null(res$plot)) {
        print(res$plot) # Prints the plot to the open PDF device
      }
    }, error = function(e){
      print(paste("SKIPPING plot for", fname, ":", e$message))
    })
  }
  
  dev.off() # Close and save the PDF
  print(paste("Done! PDF report saved successfully to:", output_pdf_path))
}

# TEST 1: PAIRED COMPARISON (1:1 Validation) #####
# Limits comparison to: DBHcore, DBH_disc, Detrended_resistance_stripped, Resistance_3cm
# TEST 1: PAIRED COMPARISON (Optimized & Robust Type-Safety)
run_test_1_paired <- function(rgp_dir, manual_excel_file, precalc_auto_file = NULL) {
  
  # --- Step 1: Get Automated Data (Fast Load or Full Process) ---
  if (!is.null(precalc_auto_file) && file.exists(precalc_auto_file)) {
    print(paste("Loading pre-calculated automated data from:", basename(precalc_auto_file)))
    
    sheets <- readxl::excel_sheets(precalc_auto_file)
    auto_list <- lapply(sheets, function(s) {
      # Read as character to prevent binding errors, then fix types later
      d <- readxl::read_excel(precalc_auto_file, sheet = s, col_types = "text")
      if(!"Source_Sheet" %in% names(d)) d$Source_Sheet <- s
      return(d)
    })
    auto_df <- bind_rows(auto_list)
    
  } else {
    print("No pre-calculated file found. Processing raw .rgp files (this may take time)...")
    auto_df <- batch_process_directory(rgp_dir)
  }
  
  # Ensure Automated Types are Clean
  auto_df <- auto_df %>%
    mutate(
      File_Clean = tools::file_path_sans_ext(File_name),
      Match_Key_Sheet = tolower(trimws(Source_Sheet)),
      across(c(DBHcore, DBH_disc, Detrended_resistance_stripped, Resistance_3cm), 
             ~suppressWarnings(as.numeric(as.character(.))))
    )
  
  # --- Step 2: Load Manual Data (All Sheets with Type Safety) ---
  print("--- Step 2: Loading Manual Data (All Sheets) ---")
  
  sheet_names <- readxl::excel_sheets(manual_excel_file)
  man_list <- list()
  
  for(sheet in sheet_names){
    tryCatch({
      # Read sheet, forcing all columns to text to avoid "Double vs Character" errors
      temp_df <- read_excel(manual_excel_file, sheet = sheet, col_types = "text")
      
      temp_df$Source_Sheet_Manual <- sheet
      
      if("File_name" %in% names(temp_df)){
        man_list[[sheet]] <- temp_df
      }
    }, error = function(e) print(paste("Skipping sheet", sheet, ":", e$message)))
  }
  
  # Now binding is safe because everything is text
  man_df <- bind_rows(man_list)
  
  # --- Step 3: Clean & Join ---
  man_df <- man_df %>%
    rename_with(~"Resistance_3cm", .cols = matches("Resistance\\(3cm\\)|Resistance.3cm", ignore.case = TRUE)) %>%
    rename_with(~"DBH_disc", .cols = matches("DBH.disc|DBH_disc", ignore.case = TRUE)) %>%
    mutate(
      File_Clean = tools::file_path_sans_ext(File_name),
      Match_Key_Sheet = tolower(trimws(Source_Sheet_Manual)),
      
      # Now safely convert text back to numbers
      DBHcore = as.numeric(as.character(DBHcore)),
      DBH_disc = as.numeric(as.character(DBH_disc)),
      Detrended_resistance_stripped = as.numeric(as.character(Detrended_resistance_stripped)),
      Resistance_3cm = as.numeric(as.character(Resistance_3cm))
    )
  
  joined_data <- inner_join(
    auto_df, 
    man_df, 
    by = c("File_Clean", "Match_Key_Sheet"),
    suffix = c(".auto", ".man")
  )
  
  print(paste("SUCCESS: Matched", nrow(joined_data), "files uniquely."))
  
  if(nrow(joined_data) == 0) {
    print("WARNING: No matches found. Check Key Examples:")
    print(head(unique(auto_df$Match_Key_Sheet)))
    return(NULL)
  }
  
  # --- Step 4: Plotting & Stats ---
  metrics <- c("DBHcore", "DBH_disc", "Detrended_resistance_stripped", "Resistance_3cm")
  plots <- list()
  
  for(m in metrics){
    col_man <- paste0(m, ".man")
    col_auto <- paste0(m, ".auto")
    if(!col_man %in% names(joined_data)) next
    
    p <- ggplot(joined_data, aes_string(x = col_man, y = col_auto)) +
      geom_point(alpha=0.5, color="blue") +
      geom_abline(intercept = 0, slope = 1, linetype="dashed", color="red", size=1) +
      labs(title = paste("1:1 Validation:", m), subtitle = paste("n =", nrow(joined_data)),
           x = "Manual", y = "Automated") + theme_minimal()
    plots[[m]] <- p
  }
  grid.arrange(grobs = plots, ncol = 2)
  
  print("--- Bias Statistics ---")
  print(joined_data %>%
          summarise(
            Bias_DBHcore = mean(DBHcore.auto - DBHcore.man, na.rm=TRUE),
            MAE_DBHcore = mean(abs(DBHcore.auto - DBHcore.man), na.rm=TRUE)
          ))
}

# TEST 2: BENCHMARKING (Optimized)
run_test_2_distribution <- function(new_trial_dir, historical_dir_path, target_files, precalc_auto_file = NULL) {
  
  # --- STEP 1: Process New Trial (Fast Load) ---
  if (!is.null(precalc_auto_file) && file.exists(precalc_auto_file)) {
    print(paste("Loading pre-calculated new trial data:", basename(precalc_auto_file)))
    sheets <- readxl::excel_sheets(precalc_auto_file)
    auto_df <- bind_rows(lapply(sheets, function(s) read_excel(precalc_auto_file, sheet = s, col_types = "text")))
  } else {
    print("Processing new trial from raw .rgp files...")
    auto_df <- batch_process_directory(new_trial_dir)
  }
  
  # Clean Auto Data
  auto_df <- auto_df %>%
    mutate(Source = "New Trial (Automated)",
           File_Clean = tools::file_path_sans_ext(File_name),
           across(c(DBHcore, DBH_disc, Detrended_resistance_stripped, Resistance_3cm), 
                  ~suppressWarnings(as.numeric(as.character(.)))))
  
  # --- STEP 2: Load Specific Manual Files ---
  print(paste("Looking for", length(target_files), "specific manual files..."))
  all_files_in_dir <- list.files(historical_dir_path, full.names = TRUE)
  manual_files <- all_files_in_dir[basename(all_files_in_dir) %in% target_files]
  
  if(length(manual_files) == 0) stop("No matching manual files found.")
  
  history_list <- list()
  
  for(mfile in manual_files){
    fname <- basename(mfile)
    
    # 2a. Exclusion Check (Prevent Self-Comparison)
    tryCatch({
      # Check overlap using first sheet (read as text for safety)
      first_sheet <- read_excel(mfile, n_max = 50, col_types = "text") %>%
        rename_with(~"DBH_disc", .cols = matches("DBH.disc|DBH_disc", ignore.case = TRUE)) %>%
        mutate(File_Clean = tools::file_path_sans_ext(File_name),
               DBH_disc = as.numeric(as.character(DBH_disc)))
      
      matches <- inner_join(first_sheet, auto_df, by = "File_Clean") %>%
        filter(abs(DBH_disc.x - DBH_disc.y) < 0.5) 
      
      if(nrow(matches) > 0.5 * nrow(first_sheet)){
        print(paste("EXCLUDING:", fname, "- is current dataset."))
        next 
      }
    }, error = function(e) { NULL }) 
    
    # 2b. Load All Sheets
    print(paste("Loading History:", fname, "..."))
    sheets <- excel_sheets(mfile)
    file_data <- lapply(sheets, function(s) {
      tryCatch({
        d <- read_excel(mfile, sheet = s, col_types = "text") %>% # Force Text
          rename_with(~"Resistance_3cm", .cols = matches("Resistance\\(3cm\\)|Resistance.3cm", ignore.case = TRUE)) %>%
          rename_with(~"DBH_disc", .cols = matches("DBH.disc|DBH_disc", ignore.case = TRUE)) %>%
          mutate(across(c(DBHcore, DBH_disc, Detrended_resistance_stripped, Resistance_3cm), 
                        ~suppressWarnings(as.numeric(as.character(.))))) %>%
          select(any_of(c("DBHcore", "DBH_disc", "Detrended_resistance_stripped", "Resistance_3cm")))
        if(nrow(d) > 0) return(d) else return(NULL)
      }, error = function(e) return(NULL))
    })
    
    valid_dfs <- Filter(Negate(is.null), file_data)
    if(length(valid_dfs) > 0) history_list[[fname]] <- bind_rows(valid_dfs)
  }
  
  if(length(history_list) == 0) return(NULL)
  
  manual_big_df <- bind_rows(history_list) %>% mutate(Source = "Historical (Manual)")
  
  # --- STEP 3: Combine & Visualize ---
  cols <- c("Source", "DBHcore", "DBH_disc", "Detrended_resistance_stripped", "Resistance_3cm")
  combined <- bind_rows(auto_df %>% select(any_of(cols)), manual_big_df %>% select(any_of(cols)))
  
  # Plots
  metrics <- c("DBHcore", "DBH_disc", "Detrended_resistance_stripped", "Resistance_3cm")
  plots <- list()
  for(m in metrics){
    if(!m %in% names(combined) || all(is.na(combined[[m]]))) next
    p <- ggplot(combined, aes_string(x = m, fill = "Source")) +
      geom_density(alpha = 0.5) + labs(title = paste("Benchmark:", m)) + theme_minimal() + theme(legend.position="bottom")
    plots[[m]] <- p
  }
  grid.arrange(grobs = plots, ncol = 2) 
  
  print("--- Benchmark Summary ---")
  print(combined %>% group_by(Source) %>% summarise(Count = n(), Mean_DBH = mean(DBHcore, na.rm=TRUE)))
}

# SECTION 6: EXECUTION BLOCK ####

# --- PATH CONFIGURATION
rgp_trial_path <- here("Moray 66 Complete") 
single_manual_file <- here("Processed data", "Moray 66.xlsx")
processed_data_folder <- here("Processed data")

# Define the file that holds your automated results
# (Note: This is the file you create in Section 7. If it exists, we read it to save time)
auto_output_file <- here("Processed Data", paste0(basename(rgp_trial_path), "_Automated_Output.xlsx"))

# --- DEFINE THE SPECIFIC FILES FOR TEST 2 ---
manual_benchmark_files <- c("Llandovery 31.xlsx", "Moray 62.xlsx", "Moray 63.xlsx", "Moray 66.xlsx")

# --- RUN TESTS

# OPTION A: Test 1 (1:1 Validation)
# Now accepts 'precalc_auto_file'. If this file exists, it skips raw processing.
print("Running Test 1: Paired Validation...")
run_test_1_paired(rgp_trial_path, single_manual_file, precalc_auto_file = auto_output_file)

# OPTION B: Test 2 (Benchmark)
print("Running Test 2: Benchmarking...")
run_test_2_distribution(rgp_trial_path, processed_data_folder, manual_benchmark_files, precalc_auto_file = auto_output_file)


