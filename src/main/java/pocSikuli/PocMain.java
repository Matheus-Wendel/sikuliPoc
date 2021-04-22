package pocSikuli;

import java.util.ArrayList;
import java.util.List;

import org.sikuli.basics.Settings;
import org.sikuli.script.FindFailed;
import org.sikuli.script.ImagePath;
import org.sikuli.script.Key;
import org.sikuli.script.Match;
import org.sikuli.script.Pattern;
import org.sikuli.script.Region;
import org.sikuli.script.Screen;
import org.sikuli.script.ScreenImage;

public class PocMain {

	public static void main(String[] args) throws FindFailed {
		ImagePath.add(System.getProperty("user.dir"));
		Settings.OcrTextRead = true;
		Settings.OcrTextRead = true;
		abreMstsc();
		abreExcel();
		salvaArquivoExcel();

	}

	public static void abreVariaveis() {
		Screen s = new Screen();

		try {
			s.click("imgs\\1618396827304.png", 0);
			s.wait("imgs\\1618403444950.png", 100);
			s.type("vari");
			Region menu = s.find("imgs\\1618398969847.png").below(300);
			Match iconVariaveis = menu.find("imgs\\1618400545429.png");
			s.click(iconVariaveis);
			s.wait("imgs/1618590987835.png", 100);
			Match find = s.find("imgs/1618590987835.png");
			Region below = find.below(440);
			// String text = below.text();
			Match findbtnVariaveis = below.find("imgs/1618788104715.png");
			s.click(findbtnVariaveis);
			// System.err.println(text);
			// https://stackoverflow.com/questions/36744457/text-recognition-is-not-working-with-sikuli-for-some-words

			ScreenImage capture = s.capture(below);

			capture.getFile("imgs", "tela variaveis");

		} catch (FindFailed e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void abreMstsc() {
		Screen s = new Screen();

		try {
			s.click("imgs/1619086491960.png", 0);

		} catch (FindFailed e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void salvaArquivoExcel() {
		Screen s = new Screen();

		try {
			Match btnArquivo = s.wait("imgs/1619087843643.png", 1000);
			s.click(btnArquivo);
			Match btnSalvarComo = s.wait("imgs/1619087991212.png", 1000);
			s.click(btnSalvarComo);
			Match btnEsteComputador = s.wait("imgs/1619088051725.png", 1000);
			s.doubleClick(btnEsteComputador);
			s.wait("imgs/1619088146892.png", 5000);
			Pattern exact = new Pattern("imgs/1619088214518.png").exact();
			Match btnDesktop = s.wait(exact, 1000);
			s.doubleClick(btnDesktop);
			s.keyDown(Key.ALT);
			s.keyDown("s");
			s.keyUp();

		} catch (FindFailed e) {
			e.printStackTrace();
		}
	}

	public static void abreExcel() {
		Screen s = new Screen();
		try {
			s.click("imgs\\1618396827304.png", 0);
			s.wait("imgs\\1618403444950.png", 1000);
			s.type("exce");
			Region menu = s.find("imgs\\1618398969847.png").below(300);
			Match iconExcel = menu.find("imgs\\1618805516813.png");
			s.click(iconExcel);
			s.wait("imgs/1618805612119.png", 5000).click();
			s.wait("imgs/1618808502527.png", 5000);
			Region menuFonteECor = s.find("imgs/1618808502527.png").above(30);

			Match botaoFontesEtamanho = menuFonteECor.find("imgs/1618808844268.png");
			// botaoFontesEtamanho.offset(-10, 0);
			s.click(botaoFontesEtamanho.offset(-20, 0));
			// Region listaDeFontes = s.wait("imgs/1618829105542.png", 5000).grow(0, 13, 0,
			// 681);
			Region listaDeFontes = s.wait("imgs/2021-04-19 08_31_46-.png", 500).below(630);

			List<Match> findLines = listaDeFontes.findLines();
			List<String> fontes = new ArrayList<String>();
			// ADICIONA PRIMEIRAS LINHAS
			for (int i = 0; i < findLines.size(); i++) {
				ScreenImage captureLine = s.capture(findLines.get(i));
				fontes.add(findLines.get(i).text());

				captureLine.getFile("imgs/linha", "l" + i);
			}

			Region fimLista = s.wait("imgs/2021-04-19 09_06_06-.png", 500).above(30);
			// ScreenImage capture = s.capture(fimLista);
			Match desceUm = fimLista.find("imgs/1618835038351.png");
			// capture.getFile("imgs", "tela variaveis");
			// Match exists = s.exists("imgs/1618915585598.png");
			// s.hover(exists);
			// https://stackoverflow.com/questions/33824784/how-to-find-exact-match-of-an-image-in-sikuli-with-java
			// Pattern exact = new Pattern("imgs/1618915585598.png").exact();
			int a = 0;
			while (s.exists("imgs/1618915585598.png") == null || a != 10) {
				s.click(desceUm);
				List<Match> ultimaLinha = fimLista.findLines();
				fontes.add(ultimaLinha.get(0).text());
				a++;

			}

			// Region listaDeFontes = s.find("imgs/1618808502527.png").grow(0, 91, 0, 681);
			// ScreenImage capture = s.capture(listaDeFontes);

			// capture.getFile("imgs", "tela variaveis");
			// listaDeFontes.setRows(21);
			// for (int i = 0; i < 21; i++) {
			// ScreenImage captureLine = s.capture(listaDeFontes.getRow(i));

			// captureLine.getFile("imgs/linha", "l" + i);
			// }
			s.keyDown(Key.ESC);
			s.keyUp();
			for (String string : fontes) {
				try {
					s.type(string);
				} catch (Exception e) {
					System.err.println(e);
				}
				s.keyDown(Key.ENTER);
				s.keyUp();
			}
			// String text = listaDeFontes.text();
			// System.err.println(text);
			// ScreenImage capture = s.capture(listaDeFontes);

			// capture.getFile("imgs", "tela variaveis");
			// Match dropDownIcon = fontSelector.find("imgs/1618807688545.png");
			// s.click(dropDownIcon);

			// System.err.println(text);
			// https://stackoverflow.com/questions/36744457/text-recognition-is-not-working-with-sikuli-for-some-words

		} catch (FindFailed e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
