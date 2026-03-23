# ========== Stage 1: Build ==========
FROM node:22-alpine AS builder

WORKDIR /app

# Install dependencies first (layer cache)
COPY package.json package-lock.json ./
RUN npm ci --ignore-scripts

# Copy source
COPY webpack.config.js ./
COPY scripts/ ./scripts/
COPY src/ ./src/
COPY manifest.xml ./

# Accept server URL as build arg, map to ENV for the build script
ARG SERVER_URL
ENV SERVER_URL=${SERVER_URL:-__SERVER_URL__}

# Build production bundle
RUN npm run build:prod

# ========== Stage 2: Serve ==========
FROM nginx:alpine

# Remove default site
RUN rm -rf /usr/share/nginx/html/*

# Copy custom nginx config
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copy built assets from builder
COPY --from=builder /app/dist/ /usr/share/nginx/html/

# Copy entrance script
COPY --from=builder /app/scripts/entrypoint.sh /usr/local/bin/entrypoint.sh
RUN chmod +x /usr/local/bin/entrypoint.sh

EXPOSE 80

ENTRYPOINT ["/usr/local/bin/entrypoint.sh"]
CMD ["nginx", "-g", "daemon off;"]
