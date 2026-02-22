FROM node:20-alpine
WORKDIR /app
COPY package.json package-lock.json* /app/
RUN npm install
COPY src /app/src
COPY index.html /app/index.html
COPY vite.config.ts /app/vite.config.ts
EXPOSE 5173
CMD ["npm", "run", "dev", "--", "--host", "0.0.0.0", "--port", "5173"]
