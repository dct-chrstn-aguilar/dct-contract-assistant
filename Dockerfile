FROM node:22-bookworm-slim AS build

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .

ARG VITE_ENTRA_CLIENT_ID
ARG VITE_ENTRA_TENANT_ID
ARG VITE_FABRIC_SCOPE
ENV VITE_ENTRA_CLIENT_ID=$VITE_ENTRA_CLIENT_ID \
    VITE_ENTRA_TENANT_ID=$VITE_ENTRA_TENANT_ID \
    VITE_FABRIC_SCOPE=$VITE_FABRIC_SCOPE

RUN npm run build

FROM node:22-bookworm-slim AS runtime

ENV NODE_ENV=production
ENV PORT=3001
WORKDIR /app

COPY package*.json ./
# The existing production entrypoint is `npm start`, which runs `tsx`.
# Keep the project's runtime tooling available so the source server starts
# exactly as it does outside Docker.
RUN npm ci --include=dev && npm cache clean --force

COPY --from=build /app/dist ./dist
COPY --from=build /app/server ./server
COPY --from=build /app/api ./api

EXPOSE 3001

CMD ["npm", "start"]
