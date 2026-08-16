FROM alpine:3.24
WORKDIR /src
COPY . .
LABEL org.opencontainers.image.source="https://github.com/mafzalkalwardev/multi-smtp-email-automation"
CMD ["sh", "-c", "echo 'multi-smtp-email-automation source package' && ls -1"]
