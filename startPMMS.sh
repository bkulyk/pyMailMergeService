#!/usr/bin/env sh
while true; do
	sleep 1 #just a puase to make sure the port is released before it starts up again
	cd /usr/share/pyMailMergeService
	/usr/bin/env python pyMailMergeService.py
done
