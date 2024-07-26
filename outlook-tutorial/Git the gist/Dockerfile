FROM node:20.16.0-alpine

WORKDIR /app

COPY package*.json ./

RUN npm install

COPY ./dist .

EXPOSE 3000

CMD ["/app/start.sh"]