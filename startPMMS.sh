#!/usr/bin/env sh
sleep 2
SCRIPT="/usr/bin/python /usr/share/pyMailMergeService/pyMailMergeService.py"
while true; do
        cd /usr/share/pyMailMergeService
        $SCRIPT  ######### The Service that runs
        sleep 1
done