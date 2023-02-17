#!/bin/bash

# prune docker
docker stop $(docker ps --filter status=running --filter name=skillhunter -q)
docker rm -f $(docker ps --filter status=exited -q)
docker rmi -f $(docker images skillhunter* -q)
docker image prune -f

# prepare new deployment folder
mv skillhunter/ oldSkillhunter/
git clone git@github.com:gnos28/skillhunter.git
cd skillhunter/
git pull -f --rebase origin main

# récupérer les .env uploadés précédemment avec scp et les déplacer ici
mv ../dotenv/.env.backend back/.env
mv ../dotenv/auth.json back/auth.json

# build docker images
docker compose -f docker-compose.prod.yml build --no-cache

# start containers
docker compose -f docker-compose.prod.yml up >~/logs/log.compose.$(date +"%s") 2>&1 &
disown

# delete old folder
sudo rm -Rf ~/oldSkillhunter/
