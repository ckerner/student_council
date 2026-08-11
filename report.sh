#!/bin/bash

POINTS_FILE="student_council_points.xlsx"

REPORT=$1

XLEAK=$( which xleak 2>/dev/null )
if [[ "x${XLEAK}" == "x" ]] ; then
   echo "\n\tERROR: xleak is not installed"
   exit 100
fi

function print_usage() 
{
   PROG=$( basename $0 )
   cat <<EOUSAGE

   Usage: ${PROG} [ Students | Events | Attendance | Leaderboard | Summary ]

EOUSAGE

}

if [[ "x${REPORT}" == "x" ]] ; then
   print_usage
   exit 1
fi

if [[ "${REPORT}" == "all" ]] ; then
   for REP in Students Events Attendance Leaderboard Summary
       do echo "Report: ${REP}"
       ${XLEAK} -s ${REP} ${POINTS_FILE} | \
           grep -E "\+|===|---|\|"
       echo " "
   done
else
   ${XLEAK} -s ${REPORT} ${POINTS_FILE} | \
      grep -E "\+|===|---|\|"
fi



