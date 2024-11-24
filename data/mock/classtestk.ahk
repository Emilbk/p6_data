    loopStartTid := A_Now
    loopSlutTid := DateAdd(loopStartTid, 6013, "Seconds")
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidTime := Floor(slutTidDifferenceSec / 60 / 60)
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    return