package oz.med.DMSParser;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.ConfigurableApplicationContext;
import oz.med.DMSParser.services.MyTrayIcon;

@SpringBootApplication
//@EnableConfigurationProperties(StorageProperties.class)
public class DmsParserApplication {

	public static ConfigurableApplicationContext context;

	public static void main(String[] args) {
//		SpringApplication.run(DmsParserApplication.class, args);

		SpringApplicationBuilder builder = new SpringApplicationBuilder(DmsParserApplication.class);
		builder.headless(false);
		context = builder.run(args);

	}

}
