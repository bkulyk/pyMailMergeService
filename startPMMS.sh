#!/usr/bin/env sh
while true; do
	sleep 1 #just a puase to make sure the port is released before it starts up again
	THEPATH="`dirname $0`";
	/usr/bin/env python $THEPATH/src/mmsd.py --no-daemon
done
