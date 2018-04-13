package io.github.plajdo.excel.update;

import java.io.File;
import java.io.InputStream;
import java.lang.ProcessBuilder.Redirect;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import javax.swing.JOptionPane;

public class Updater{
	
	public static void main(String[] args){
		try{
			start();
		}catch(Exception e){
			JOptionPane.showMessageDialog(null, "Chyba pri sp\u00FA\u0161\u0165an\u00ED! Popis chyby:\n" + e.toString(), "Chyba", JOptionPane.ERROR_MESSAGE);
			e.printStackTrace();
			System.exit(-1);
		}
		
		System.exit(0);
		
	}
	
	public static void start() throws Exception{
		String jvm_location;
		Process p;
		
		AutoUpdate.updateData();
		URL web = new URL(AutoUpdate.GithubData.getAssets_download_url());
		Path out = Paths.get(System.getProperty("java.io.tmpdir") + File.separator + "excelstuff_" + AutoUpdate.GithubData.getRelease_name() + ".jar");
		
		File outFile = out.toFile();
		outFile.deleteOnExit();
		
		try(InputStream in = web.openStream()){
			Files.copy(in, out, StandardCopyOption.REPLACE_EXISTING);
		}
		
		if(System.getProperty("os.name").startsWith("Win")){
		    jvm_location = System.getProperties().getProperty("java.home") + File.separator + "bin" + File.separator + "java.exe";
		}else{
		    jvm_location = System.getProperties().getProperty("java.home") + File.separator + "bin" + File.separator + "java";
		}
		
		ProcessBuilder pb = new ProcessBuilder(jvm_location, "-jar", outFile.toString());
		pb.redirectOutput(Redirect.INHERIT);
		pb.redirectError(Redirect.PIPE);
		
		p = pb.start();
		p.waitFor();
		
	}
	
}
