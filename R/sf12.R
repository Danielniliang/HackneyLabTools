#' SF-12 Scoring Function
#'
#' Calculates SF-12 subscale scores, z-scores, norm-based scores, and composite summary scores (PCS and MCS).
#'
#' @param data A data frame containing the 12 SF-12 item responses. Items should be ordered according to the SF-12 standard format.
#'
#' @return A data frame with raw scores, transformed scores, z-scores, norm-based scores, and composite scores:
#' \itemize{
#'   \item Raw scores for each domain (PF, RP, BP, GH, VT, SF, RE, MH)
#'   \item Transformed scores scaled to 0–100
#'   \item Z-scores and norm-based scores for each domain
#'   \item Composite scores: SF12_PCS (Physical Component Summary), SF12_MCS (Mental Component Summary)
#' }
#' @export
SF12 <- function(data) {
  data <- as.data.frame(apply(data, 2, as.numeric))
  sf12 <- c()
  sf12$PF <- data[,2] + data[,3]
  sf12$RP <- data[,4] + data[,5]
  sf12$BP <- data[,8]
  sf12$GH <- data[,1]
  sf12$VT <- data[,10]
  sf12$SF <- data[,12]
  sf12$RE <- data[,6] + data[,7]
  sf12$MH <- data[,9] + data[,11]
  sf12 <- as.data.frame(sf12)

  sf12$TSPF <- (sf12$PF - 2)/4 * 100
  for (i in 2:(ncol(sf12)-1)) {
    if (i %in% c(3,4,5,6)) {
      sf12[,i+8] <- (sf12[,i] - 1)/4 * 100
    } else {
      sf12[,i+8] <- (sf12[,i] - 2)/8 * 100
    }
  }

  sf12[,17] <- (sf12[,9] - 81.18122)/29.10558
  sf12[,18] <- (sf12[,10] - 80.52856)/27.13526
  sf12[,19] <- (sf12[,11] - 81.74015)/24.53019
  sf12[,20] <- (sf12[,12] - 72.19795)/23.19041
  sf12[,21] <- (sf12[,13] - 55.5909)/24.8438
  sf12[,22] <- (sf12[,14] - 83.73973)/24.75775
  sf12[,23] <- (sf12[,15] - 86.41051)/22.35543
  sf12[,24] <- (sf12[,16] - 70.18217)/20.50597

  for (i in 17:24) {
    sf12[,i+8] <- (sf12[,i] * 10) + 50
  }

  sf12[,33] <- (sf12[,17]*0.42402) + (sf12[,18]*0.35119) + (sf12[,19]*0.31754) + (sf12[,20]*0.24954) +
               (sf12[,21]*0.02877) + (sf12[,22]*-0.00753) + (sf12[,23]*-0.19206) + (sf12[,24]*-0.22069)
  sf12[,34] <- (sf12[,17]*-0.22999) + (sf12[,18]*-0.12329) + (sf12[,19]*-0.09731) + (sf12[,20]*-0.01571) +
               (sf12[,21]*0.23534) + (sf12[,22]*0.26876) + (sf12[,23]*0.43407) + (sf12[,24]*0.48581)

  sf12[,35] <- 50 + (sf12[,33] * 10)
  sf12[,36] <- 50 + (sf12[,34] * 10)

  colnames(sf12) <- c(
    'raw score Physical Functioning (PF)', 'raw score Role Physical (RP)', 'raw score Bodily Pain (BP)',
    'raw score General Health (GH)', 'raw score Vitality (VT)', 'raw sore Social Functioning (SF)',
    'raw score Role Emotional (RE)', 'raw score Mental Health (MH)', 'transformed score PF', 'transformed score RP',
    'transformed score BP', 'transformed score GH', 'transformed score VT', 'transformed score SF',
    'transformed score RE', 'transformed score MH', 'PF z-score', 'RP z-score', 'BP z-score', 'GH z-score',
    'VT z-score', 'SF z-score', 'RE z-score', 'MH z-score', 'PF norm based', 'RP norm based ', 'BP norm based ',
    'GH norm based ', 'VT norm based ', 'SF norm based ', 'RE norm based ', 'MH norm based ',
    'Aggregate PHYSICAL', 'Aggregate MENTAL', 'SF12_PCS', 'SF12_MCS'
  )

  return(sf12)
}
