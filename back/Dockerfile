FROM node:18 as base

WORKDIR /back
COPY package*.json ./
EXPOSE 5000

FROM base as production
RUN npm install
COPY ./src /back/src
COPY .env /back/
COPY auth.json /back/
COPY tsconfig.json /back/
COPY tsconfig.build.json /back/
RUN npm install -g typescript ts-node
RUN npm install -g pm2
RUN npm run build
CMD ["ts-node", "dist/index.js"]

FROM base as dev
ENV NODE_ENV=development
RUN npm install -g nodemon && npm install
COPY ./src /back/src
COPY .env /back/
COPY auth.json /back/
COPY tsconfig.json /back/
COPY tsconfig.build.json /back/
CMD ["nodemon", "src/index.js"]

# WORKDIR /src
# ENV NODE_ENV=production
# RUN npm ci
# RUN npm i -D typescript ts-node @types/node @types/express
