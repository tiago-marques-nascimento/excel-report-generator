package com.tiago;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import com.tiago.datasource.DataSource;
import com.tiago.datasource.Student;
import com.tiago.datasource.Teacher;
import com.tiago.datasource.Class;
import com.tiago.datasource.ClassTypeEnum;
import com.tiago.utils.SpreadSheetWriter;

public class Main {

	public static void main(String args[]){

		Student studentTiago = new Student(
			"Tiago",
			32,
			"Rua Antonio Francisco, 672",
			"(32) 992736726",
			"tiago@gmail.com",
			new ArrayList<Class>()
		);

		Student studentAnderson = new Student(
			"Anderson",
			23,
			"Rua Antonio Francisco, 572",
			"(32) 992711724",
			"anderson@gmail.com",
			new ArrayList<Class>()
		);

		Student studentEduardo = new Student(
			"Eduardo",
			41,
			"Rua Antonio Francisco, 272",
			"(32) 992736726",
			"eduardo@gmail.com",
			new ArrayList<Class>()
		);

		Student studentRafael = new Student(
			"Rafael",
			27,
			"Rua Antonio Francisco, 272",
			"(32) 933336722",
			"rafael@gmail.com",
			new ArrayList<Class>()
		);

		Student studentTaynara = new Student(
			"Taynara",
			25,
			"Rua Antonio Francisco, 172",
			"(32) 992111116",
			"taynara@gmail.com",
			new ArrayList<Class>()
		);

		Teacher teacherDiogo = new Teacher(
			"Diogo",
			32,
			"Rua Antonio Francisco, 672",
			"(32) 982736725",
			"diogo@gmail.com",
			new ArrayList<Class>()
		);

		Teacher teacherLivia = new Teacher(
			"Livia",
			52,
			"Rua Antonio Francisco, 972",
			"(32) 982799925",
			"livia@gmail.com",
			new ArrayList<Class>()
		);

		Teacher teacherMaria = new Teacher(
			"Maria",
			29,
			"Rua Fernando Francisco, 272",
			"(32) 911136711",
			"maria@gmail.com",
			new ArrayList<Class>()
		);

		Class mathClass = new Class(
			"Math",
			ClassTypeEnum.DAY_CLASS,
			4.5,
			List.of(
				studentTiago,
				studentRafael,
				studentTaynara
			),
			teacherDiogo);
		teacherLivia.getClasses().add(mathClass);
		studentTiago.getClasses().add(mathClass);
		studentRafael.getClasses().add(mathClass);
		studentTaynara.getClasses().add(mathClass);

		Class biologyClass = new Class(
			"Biology",
			ClassTypeEnum.NIGHT_CLASS,
			3.2,
			List.of(
				studentTiago,
				studentAnderson,
				studentEduardo
			),
			teacherLivia);
		teacherLivia.getClasses().add(biologyClass);
		studentTiago.getClasses().add(biologyClass);
		studentAnderson.getClasses().add(biologyClass);
		studentEduardo.getClasses().add(biologyClass);

		Class chemistryClass = new Class(
			"Chemistry",
			ClassTypeEnum.DAY_CLASS,
			4.0,
			List.of(
				studentAnderson,
				studentEduardo,
				studentRafael,
				studentTaynara
			),
			teacherMaria);
		teacherLivia.getClasses().add(chemistryClass);
		studentAnderson.getClasses().add(chemistryClass);
		studentEduardo.getClasses().add(chemistryClass);
		studentRafael.getClasses().add(chemistryClass);
		studentTaynara.getClasses().add(chemistryClass);

		DataSource dataSource = new DataSource(
			"Students Records",
			List.of(
				studentTiago,
				studentAnderson,
				studentEduardo,
				studentRafael,
				studentTaynara
			)
		);

		SpreadSheetWriter.write(
			Optional.of("There's no data"), "students-template.xlsx", dataSource);

		System.out.println("Report generated.");
	}
}
