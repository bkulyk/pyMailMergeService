#!/usr/bin/env sh
while true; do
	sleep 1 #just a puase to make sure the port is released before it starts up again
	PATH="`dirname $0`";
	/usr/bin/env python $PATH/src/mmsd.py --no-daemon
done
