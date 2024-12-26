# Use uma imagem base do Java 17
FROM openjdk:17-slim

# Adicione um diretório para a aplicação
WORKDIR /app

# Copie o JAR gerado para a imagem
COPY target/converter-0.0.1-SNAPSHOT.jar app.jar

# Exponha a porta que sua aplicação usa
EXPOSE 8080

# Comando para executar o JAR
CMD ["java", "-jar", "app.jar"]