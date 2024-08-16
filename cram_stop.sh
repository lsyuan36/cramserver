#!/bin/bash
cd server
if [ -f server.pid ]; then
  PID=$(cat server.pid1)
  echo "Stopping server with PID $PID"
  kill $PID
  rm server.pid
  echo "Server stopped"
else
  echo "Server is not running"
fi
