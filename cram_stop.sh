#!/bin/bash

if [ -f server1.pid ]; then
  PID=$(cat server1.pid)
  echo "Stopping server with PID $PID"
  kill $PID
  rm server1.pid
  echo "Server stopped"
else
  echo "Server is not running"
fi
