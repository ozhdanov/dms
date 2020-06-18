package oz.med.DMSParser;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class DmsParserApplication {

	public static ConfigurableApplicationContext context;

	public static void main(String[] args) {
		SpringApplicationBuilder builder = new SpringApplicationBuilder(DmsParserApplication.class);
		builder.headless(false);
		context = builder.run(args);

	}

}
