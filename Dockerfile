FROM php:8.2-apache

# Install system dependencies and PHP extensions
RUN apt-get update && apt-get install -y \
    libzip-dev \
    unzip \
    sqlite3 \
    libsqlite3-dev \
    libpng-dev \
    libjpeg-dev \
    libfreetype6-dev \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install zip pdo pdo_sqlite gd \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Increase PHP upload and memory limits for large Excel sheets
RUN echo "upload_max_filesize = 64M\npost_max_size = 64M\nmemory_limit = 512M\nmax_execution_time = 300" > /usr/local/etc/php/conf.d/custom.ini

# Enable Apache mod_rewrite
RUN a2enmod rewrite

# Set working directory
WORKDIR /var/www/html

# Copy application source code
COPY src/ /var/www/html/

# Unpack vendor dependencies if supplied as archive
RUN if [ ! -d /var/www/html/vendor ] && [ -f /var/www/html/vendor.tar.gz ]; then \
        mkdir -p /var/www/html/vendor && \
        tar -xzf /var/www/html/vendor.tar.gz -C /var/www/html/vendor && \
        rm -f /var/www/html/vendor.tar.gz; \
    fi

# Create persistent storage directories with proper permissions for Apache (www-data)
RUN mkdir -p /var/warehouse_data/packing_lists /var/warehouse_data/archive \
    && chown -R www-data:www-data /var/warehouse_data /var/www/html \
    && chmod -R 775 /var/warehouse_data

EXPOSE 80
