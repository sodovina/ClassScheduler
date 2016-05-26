package classScheduler_src;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

public class ConfigParser {
	private ArrayList<Subjects> subjects = new ArrayList<Subjects>();
	private ArrayList<Classes> rooms = new ArrayList<Classes>();

	public ConfigParser(String path) {
		getInputs(path);
	}

	private void getInputs(String path) {
		try {
			BufferedReader br = new BufferedReader(new FileReader(path));
			String line;
			while ((line = br.readLine()) != null) {
				if (line.contains("#Lendet")) {
					parseSubjects(br);
				} else if (line.contains("#Klaset")) {
					parseRooms(br);
				}
			}
			br.close();
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}

	private void parseSubjects(BufferedReader b) throws IOException {
		String line, subject = "", prof = "";
		int department = 0, year = 0, students = 0;
		boolean lab = false;
		while (!(line = b.readLine()).contains("#end")) {
			if (line.toLowerCase().startsWith("drejtimi")) {
				department = new Integer(line.substring(line.indexOf("=") + 1).trim()).intValue();
			} else if (line.toLowerCase().startsWith("viti")) {
				year = new Integer(line.substring(line.indexOf("=") + 1).trim()).intValue();
			} else if (line.toLowerCase().startsWith("lenda")) {
				subject = line.substring(line.indexOf("=") + 1).trim();
			} else if (line.toLowerCase().startsWith("prof")) {
				prof = line.substring(line.indexOf("=") + 1).trim();
			} else if (line.toLowerCase().startsWith("numri_studentave")) {
				students = new Integer(line.substring(line.indexOf("=") + 1).trim()).intValue();
			} else if (line.toLowerCase().startsWith("lab")) {
				if (line.contains("true"))
					lab = true;
			}
		}
		// System.out.printf("%d %d %s %s %d %b %n", department, year, subject,
		// prof, students, lab);
		subjects.add(new Subjects(department, year, subject, prof, students, lab));
	}

	public void parseRooms(BufferedReader b) throws IOException {
		String line, roomname = "";
		int capacity = 0;
		boolean lab = false;
		while (!(line = b.readLine()).contains("#end")) {
			if (line.toLowerCase().startsWith("emertimi")) {
				roomname = line.substring(line.indexOf("=") + 1).trim();
			} else if (line.toLowerCase().startsWith("kapaciteti")) {
				capacity = new Integer(line.substring(line.indexOf("=") + 1).trim()).intValue();
			} else if (line.toLowerCase().startsWith("lab")) {
				if (line.contains("true"))
					lab = true;
			}
		}
		// System.out.printf("%s %d %b %n", roomname, capacity, lab);
		rooms.add(new Classes(roomname, capacity, lab));
	}

	public ArrayList<Subjects> getSubjects() {
		return subjects;
	}

	public int getSubjectsCount() {
		return subjects.size();
	}

	public ArrayList<Classes> getRooms() {
		ArrayList<Classes> temp = new ArrayList<Classes>();
		for (int i = 0; i < rooms.size(); i++)
			temp.add(new Classes(rooms.get(i).getClassName(), rooms.get(i).getClassCapacity(), rooms.get(i).isLab()));
		return temp;
	}
}
