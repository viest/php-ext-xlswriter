# Docker 多阶构建

## Dockerfile

```dockerfile
FROM alpine:3.9.6 as build

# 构建xlswriter扩展，根据自身需要替换版本号
ENV XLSWRITER_VERSION 1.3.4.1

RUN apk update \
    && apk add --no-cache php7-pear php7-dev zlib-dev re2c gcc g++ make curl \
    && curl -fsSL "https://pecl.php.net/get/xlswriter-${XLSWRITER_VERSION}.tgz" -o xlswriter.tgz \
    && mkdir -p /tmp/xlswriter \
    && tar -xf xlswriter.tgz -C /tmp/xlswriter --strip-components=1 \
    && rm xlswriter.tgz \
    && cd /tmp/xlswriter \
    && phpize && ./configure --enable-reader && make && make install

#-------------------------------------------------------------------------------------------

FROM alpine:3.9.6

# 根据自身需要，添加其它软件
RUN apk update && apk add --no-cache php

COPY --from=build /usr/lib/php7/modules/xlswriter.so /usr/lib/php7/modules/xlswriter.so

RUN echo "extension=xlswriter.so" > /etc/php7/conf.d/xlswriter.ini
```

## 构建

```bash
docker build -f Dockerfile -t viest/xlswriter:1.3.4.1 .
```