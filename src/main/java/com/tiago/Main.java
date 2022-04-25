package com.tiago;

import java.util.List;

import com.tiago.datasource.DataSource;
import com.tiago.datasource.Programmer;
import com.tiago.datasource.ProgrammingLanguage;
import com.tiago.utils.SpreadSheetWriter;

public class Main {

	public static void main(String args[]){

		DataSource dataSource = new DataSource();

		Programmer tiago = new Programmer();
		tiago.setName("Tiago");
		tiago.setAge(32);
		tiago.setYearsExperience(7);

		Programmer diogo = new Programmer();
		diogo.setName("Diogo");
		diogo.setAge(33);
		diogo.setYearsExperience(2);

		ProgrammingLanguage java = new ProgrammingLanguage();
		java.setLanguage("Java");

		ProgrammingLanguage dotnet = new ProgrammingLanguage();
		dotnet.setLanguage("DotNet");

		ProgrammingLanguage python = new ProgrammingLanguage();
		python.setLanguage("Python");

		tiago.setProgrammingLanguages(List.of(
			python, java, dotnet, python, java, dotnet
		));

		diogo.setProgrammingLanguages(List.of(
			dotnet, python, java
		));

		dataSource.setProgrammers(List.of(tiago, diogo));

		SpreadSheetWriter spreadSheetWriter = new SpreadSheetWriter();
		spreadSheetWriter.write("hello-world-template.xlsx", dataSource);
		System.out.println("Hello World!");
	}
}
