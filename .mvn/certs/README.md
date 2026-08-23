# Corporate TLS-inspection trust store

This machine's network re-signs some outbound HTTPS traffic with an internal CA
(`ssl.nice.com`, chaining up to `NICE Issuing CA01` → `Nice Systems RootCA`). Windows already
trusts that root (deployed via GPO), but the JDK's own `cacerts` truststore doesn't - so plain
Java HTTPS clients (like the one Spring AI uses to download the ONNX embedding model from
Hugging Face) fail with `PKIX path building failed`, even though `curl`/browsers on the same
machine work fine.

Only `nice-systems-root-ca.pem` and `nice-issuing-ca01.pem` are committed here - they're public
certificates, not secrets. `cacerts-with-corp-ca.jks` (the actual merged truststore used at
runtime, via `pom.xml`'s `spring-boot-maven-plugin` `jvmArguments`) is machine/JDK-specific and
**not** committed (see `.gitignore`) - regenerate it once per machine:

```bash
JAVA_HOME="/c/Program Files/Java/jdk-17"   # adjust to your JDK 17+ install
cp "$JAVA_HOME/lib/security/cacerts" .mvn/certs/cacerts-with-corp-ca.jks
"$JAVA_HOME/bin/keytool" -importcert -alias nice-issuing-ca01 \
  -file .mvn/certs/nice-issuing-ca01.pem \
  -keystore .mvn/certs/cacerts-with-corp-ca.jks -storepass changeit -noprompt
"$JAVA_HOME/bin/keytool" -importcert -alias nice-systems-root-ca \
  -file .mvn/certs/nice-systems-root-ca.pem \
  -keystore .mvn/certs/cacerts-with-corp-ca.jks -storepass changeit -noprompt
```

If your machine's default `cacerts` isn't JKS format, drop the `-Djavax.net.ssl.trustStoreType=JKS`
JVM argument in `pom.xml` (or set it to `PKCS12`) to match.

If you're **not** on this corporate network (no TLS interception), none of this is needed - just
remove the `jvmArguments` block from `pom.xml` and the default JDK trust store will work fine.
