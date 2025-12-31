plugins {
    `java-library`
    `maven-publish`
    signing
}

// Project metadata from gradle.properties
group = property("group") as String
version = property("version") as String

val projectName: String by project
val projectDescription: String by project
val projectUrl: String by project
val projectScmUrl: String by project
val licenseName: String by project
val licenseUrl: String by project
val developerName: String by project
val developerEmail: String by project
val developerOrganization: String by project
val developerOrganizationUrl: String by project

// Dependency versions
val vertxVersion: String by project
val keelCoreVersion: String by project
val keelTestVersion: String by project
val excelStreamingReaderVersion: String by project
val commonsLang3Version: String by project

repositories {
    // Internal Nexus repository for dependencies
    maven {
        name = "InternalNexus"
        url = uri(findProperty("internalNexusPublicUrl") as String)
        credentials {
            username = findProperty("internalNexusUsername") as String
            password = findProperty("internalNexusPassword") as String
        }
        // Allow insecure protocol if needed (not recommended for production)
        // isAllowInsecureProtocol = false
    }

    mavenCentral()
}

dependencies {
    // Core Keel dependencies (from pom.xml)
    api("io.github.sinri:keel-core:$keelCoreVersion")

    // Apache Commons (from pom.xml - fixes vulnerability in POI)
    api("org.apache.commons:commons-lang3:$commonsLang3Version")

    // Excel streaming reader (from pom.xml)
    api("com.github.pjfanning:excel-streaming-reader:$excelStreamingReaderVersion")

    // Test dependencies (from pom.xml)
    testImplementation("io.github.sinri:keel-test:$keelTestVersion")
    testImplementation("org.junit.jupiter:junit-jupiter:5.11.4")
    testRuntimeOnly("org.junit.platform:junit-platform-launcher")
}

java {
    toolchain {
        languageVersion.set(JavaLanguageVersion.of(17))
    }
    withSourcesJar()
    withJavadocJar()
}

tasks.compileJava {
    options.encoding = "UTF-8"
    options.release.set(17)
    // Gradle will automatically compile Java modules
}

tasks.compileTestJava {
    options.encoding = "UTF-8"
    options.release.set(17)
}

// Configure resources (exclude config.properties like Maven)
tasks.processResources {
    exclude("config.properties")
}

// Configure test task (matching Maven surefire configuration)
tasks.test {
    useJUnitPlatform()
    include("io/github/sinri/keel/integration/poi/**/*Test.class")
}

// Configure JavaDoc (matching Maven javadoc plugin configuration)
tasks.javadoc {
    options.encoding = "UTF-8"
    if (options is StandardJavadocDocletOptions) {
        val stdOptions = options as StandardJavadocDocletOptions
        stdOptions.charSet = "UTF-8"
        stdOptions.docEncoding = "UTF-8"
        stdOptions.memberLevel = JavadocMemberLevel.PROTECTED
        stdOptions.docTitle = "$projectName $version Document"
        stdOptions.windowTitle = "$projectName $version Document"
        stdOptions.addBooleanOption("html5", true)
        stdOptions.addStringOption("Xdoclint:-missing", "-quiet") // 提示缺失的注释
    }
}

// Publishing configuration
publishing {
    publications {
        create<MavenPublication>("mavenJava") {
            from(components["java"])

            pom {
                name.set(projectName)
                description.set(projectDescription)
                url.set(projectUrl)

                licenses {
                    license {
                        name.set(licenseName)
                        url.set(licenseUrl)
                    }
                }

                developers {
                    developer {
                        name.set(developerName)
                        email.set(developerEmail)
                        organization.set(developerOrganization)
                        organizationUrl.set(developerOrganizationUrl)
                    }
                }

                scm {
                    url.set(projectScmUrl)
                }
            }
        }
    }

    repositories {
        // Internal Nexus repositories
        maven {
            name = "Internal"
            url = if (version.toString().endsWith("SNAPSHOT")) {
                uri(findProperty("internalNexusSnapshotsUrl") as String)
            } else {
                uri(findProperty("internalNexusReleasesUrl") as String)
            }
            credentials {
                username = findProperty("internalNexusUsername") as String
                password = findProperty("internalNexusPassword") as String
            }
        }

        // Maven Central (OSSRH)
        maven {
            name = "Release"

            if (version.toString().endsWith("SNAPSHOT")) {
                url = uri(findProperty("internalNexusSnapshotsUrl") as String)
                credentials {
                    username = findProperty("internalNexusUsername") as String
                    password = findProperty("internalNexusPassword") as String
                }
            } else {
                url = uri("https://s01.oss.sonatype.org/service/local/staging/deploy/maven2/")
                credentials {
                    username = findProperty("ossrhUsername") as String? ?: System.getenv("OSSRH_USERNAME")
                    password = findProperty("ossrhPassword") as String? ?: System.getenv("OSSRH_PASSWORD")
                }
            }

        }
    }
}

// Signing configuration
signing {
    // Only sign if not a SNAPSHOT and signing credentials are available
    setRequired({
        !version.toString().endsWith("SNAPSHOT") && gradle.taskGraph.hasTask("publish")
    })
    sign(publishing.publications["mavenJava"])
}

// Custom tasks for publishing to specific repositories
tasks.register("publishToInternal") {
    group = "publishing"
    description = "Publish to internal Nexus repository"
    dependsOn("publishMavenJavaPublicationToInternalRepository")
}

tasks.register("publishToRelease") {
    group = "publishing"
    description = "Publish to Maven Central (OSSRH)"
    dependsOn("publishMavenJavaPublicationToReleaseRepository")
}

