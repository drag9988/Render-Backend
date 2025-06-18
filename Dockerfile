FROM node:18

RUN apt-get update && \
    apt-get install -y libreoffice ghostscript && \
    apt-get clean

WORKDIR /app

COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build

EXPOSE 3000
CMD ["npm", "run", "start:prod"]