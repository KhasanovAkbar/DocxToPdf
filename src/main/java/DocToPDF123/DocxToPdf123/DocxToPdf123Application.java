package DocToPDF123.DocxToPdf123;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DocxToPdf123Application {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(DocxToPdf123Application.class, args);

		Demo demo = new Demo();

		demo.convertToXslFo();
	}



}
