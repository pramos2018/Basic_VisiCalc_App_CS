using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace NS_BasicVisiCalcApp
{
    public partial class BasicVisiCalcApp : Form
    {

	    // Project Variables

	    String currPath = "c:\\temp\\vcalc\\";
	    String currFile = "vcalc1.txt";
	    static int NROW = 100;
	    static int NCOL = 50;

        static cell[,] tb = new cell[NROW, NCOL];// Table
	    static int[] col_size = new int[NCOL];// default size
	    static int[] col_hide = new int[NCOL];// default unhide
	    static int[] row_hide = new int[NROW + 1];// default unhide
	    static int lastRow = 0;
	    static int lastCol = 0;
	    static int curCols = 0;
	    static int pr = 0, pc = 0;// For Scrolling
	    static char k;// keyboard input
	    static int prtX = 0;
	    static int prtY = 0;

        public BasicVisiCalcApp()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            //Start
            newSpreadsheet();
            preSet();
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void createTable()
        {
            int r, c, c1, r1, i, j, sz;
            int rows = 100;
            int cols = 50;

            //dbGridTable.Dispose();
            //dbGridTable = new DataGridView();

            //--------------------Resetting Table -----------------
            while (dbGridTable.Columns.Count > 0)// Removing Columns
            { dbGridTable.Columns.RemoveAt(0); }

            while (dbGridTable.Rows.Count > 0)// Removing Rows
            { dbGridTable.Rows.RemoveAt(0); }
            //------------- Adding and Setting the Columns --------
            c1 = 0;
            for (c = 1; c < cols; c++)
            {
                if (col_hide[c] == 0)
                {
                    dbGridTable.Columns.Add(getCol(c), getCol(c));
                    sz = col_size[c];
                    dbGridTable.Columns[c1].Width = sz * 6;
                    c1++;
                }
            }
            //-------  Adding the Rows and Setting the data --------

            for (r = 1; r < rows; r++)
            {
                
                String[] rw = new String[cols];
                c1 =0;
                for (c = 1; c < cols; c++)
                {
                    if (col_hide[c]==0) {
                        rw[c1] = printCell(tb[r, c], c);
                        c1++;
                    }
                }
                if (row_hide[r]==0) {
                    dbGridTable.Rows.Add(rw);
                }
            }
            //------------------Setting the Headers ---------------
            r = 1;
             foreach (DataGridViewRow row in dbGridTable.Rows)
            {
                    if (row.IsNewRow) continue;
                    row.HeaderCell.Value = r.ToString().Replace(" ","") ;
                    row.Height = 20;
                    r++;
            }

             /*c = 1;
             foreach (DataGridViewColumn col in dbGridTable.Columns)
             {
                 sz = col_size[c];
                 dbGridTable.Columns[c-1].Width = sz * 6;
                 c++;
             }
            //*/
            //dbGridTable.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //-------------------------------------------------------

        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            updateCell();
        }

        static String getCol(int c)
        {
            String str = "";
            Double d;
            Char ch;
            int c1, c2;
            c2 = c;
            if (c >= 27)
            {
                d = c / 26;
                c1 = (int) (Math.Round(d , 0));
                ch = (char) (64 + c1);
                str = str + ch.ToString();
                c2 = c - c1 * 26;
            }
            ch = (char) (64 + c2);
            str = str + ch.ToString();

            return str;
        }
        //****************** ENTIRE JAVA CODE *********************
        void preSet()
        {
            int r = 1, c = 1, i = 0;
            for (i = 1; i < 50; i++)
                col_size[i] = 8;// Default size;

            col_size[1] = 20;
            for (i = 2; i <= 8; i++)
                col_size[i] = 10;

            tb[1,1] = setCell(2, 320, 0, "Income Statement");
            tb[2,1] = setCell(2, 120, 0, "Net Sales");
            tb[3,1] = setCell(2, 120, 0, "Costs");
            tb[4,1] = setCell(2, 120, 0, "Margin");
            tb[5,1] = setCell(2, 120, 0, "----------------");
            tb[6,1] = setCell(2, 120, 0, "SG&A");
            tb[7,1] = setCell(2, 120, 0, "----------------");
            tb[8,1] = setCell(2, 120, 0, "Operating Profit");

            setTest("Jan", 2);
            setTest("Feb", 3);
            setTest("Mar", 4);
            setTest("Apr", 5);
            setTest("May", 6);
            setTest("Jun", 7);
            setTest("Jul", 8);
            setTest("Aug", 9);
            setTest("Sep", 10);
            setTest("Oct", 11);
            setTest("Nov", 12);
            setTest("Dec", 13);

            showMatrix();
        }

        void setTest(String title, int c)
        {

            tb[1, c] = setCell(2, 320, 0, title);
            tb[2, c] = setCell(1, 220, 100, "");
            tb[3, c] = setCell(1, 220, 80, "");
            tb[4, c] = setCell(1, 220, 20, "");
            tb[5, c] = setCell(2, 220, 0, "-------");
            tb[6, c] = setCell(1, 220, 10, "");
            tb[7, c] = setCell(2, 220, 0, "-------");
            tb[8, c] = setCell(1, 220, 10, "");
            int r = 8;
            if (c > lastCol)
            {
                lastCol = c;
            }
            if (r > lastRow)
            {
                lastRow = r;
            }
        }

        public void showMatrix() {
            calculateCells();
            createTable();
    	}

        void newSpreadsheet()
        {
            int r, c;
            pc = 0; pr = 0;

            for (c = 0; c < NCOL; c++)
            {
                col_hide[c] = 0;// Default
                col_size[c] = 8;// Default
            }
            for (r = 0; r < NROW; r++)
            {// reseting columns
                row_hide[r] = 0;
                for (c = 0; c < NCOL; c++)
                {
                    // delCell(r,c);
                    tb[r, c] = setCell(0, 0, 0, "");
                }
            }
            col_size[0] = 4;
            showMatrix();
        }

        // ********************** Basic Cell Operations - Set/Del/Copy **********************
        static cell setCell(int type, int Format, float value, String Text)
        {
            cell tmp = new cell();
            tmp.flag = type;// Flag: 0 - Empty, 1 - Number, 2 = Text
            tmp.format = Format;// Float 255 options - Alignment(C = Center,
            // R=Right, L=Left);
            tmp.value = value;// value
            tmp.formula = Text;// Text/Formula
            return tmp;
        }

        public void delCell(int r, int c)
        {
            if (r >= NROW || c >= NCOL || r < 0 || c < 0)
            {
                return;
            }
            try
            {
                tb[r,c].flag = 0;
                tb[r,c].formula = " ";
                tb[r,c].value = 0;
                tb[r,c].format = 0;
            }
            catch (Exception ex)
            {
            }
        }

        void copyCell(int rd, int cd, int ro, int co)
        {
            if (rd > NROW || ro > NROW || cd > NCOL || co > NCOL)
            {
                return;
            }
            try
            {
                col_size[cd] = col_size[co];
                col_hide[cd] = col_hide[cd];
                row_hide[rd] = row_hide[ro];

                tb[rd, cd].flag = tb[ro, co].flag;
                tb[rd, cd].format = tb[ro, co].format;
                tb[rd, cd].value = tb[ro, co].value;
                //tb[rd, cd].formula = tb[ro, co].formula;
                copyFormula(rd, cd, ro, co);

                if (rd > lastRow)
                {
                    lastRow = rd;
                }
                if (cd > lastCol)
                {
                    lastCol = cd;
                }
            }
            catch (Exception ex)
            {
            }

        }

        static String getCell(int r, int c)
        {
            String txt = "";
            String cl = getCol(c);
            String rw = r.ToString();
            txt = cl + rw;
            return txt;
        }

        static float getCellVal(String txt1)
        {
            float val = 0;
            char[] txt = txt1.ToCharArray();
            

            if (txt[0] >= 48 && txt[0] <= 57)
            {
                val = float.Parse(txt1);
            }
            else
            {
                RC t1 = getRC(txt1);// Cell

                if (t1.r != 0 && t1.c != 0)
                {
                    val = tb[t1.r, t1.c].value;
                }
                else
                {
                    val = 0;
                }
            }
            return val;
        }

        static RC getRC(String txt)
        {
            String x1 = "", tmp1 = txt.ToUpper().Replace(" ", "");
            

            String t_c = "";
            String t_r = "";

            RC t = new RC(0, 0);
            int i, i1 = 0, r = 0, c = 0, p = 1;

            int len = tmp1.Length;

            for (i = 0; i < len; i++)
            {
                x1 = tmp1.Substring(i, 1);
                char[] ch = x1.ToCharArray();
                if (ch[0] >= 65 && ch[0] <= 65 + 25)
                {
                    t_c = t_c + ch[0].ToString();
                    i1++;
                }
                if (ch[0] >= 48 && ch[0] <= 57)
                {
                    t_r = t_r + ch[0].ToString();
                }
            }

            char[] tc = t_c.ToCharArray();
            for (i = 0; i < i1; i++)
            {// column
                if (tc[i] == 0)
                {
                    break;
                }

                p = 26 * (i1 - i - 1);
                if (p == 0)
                {
                    p = 1;
                }

                c = c + (tc[i] - 64) * p;// *(pow(10,i));
            }

            t.c = c;
            if (t_r.Length > 0)
            {
                t.r = Convert.ToInt32(t_r);
                
            }
            else
            {
                t.r = 0;
            }
            return t;
        }

        // *************************[PRINT AND FORMAT CELLS]************************
        static String printCell(cell c, int col) {

		int size = col_size[col];
		if (size == 0) {
			return "";
		}

		char[] prt = new char[50];
		String str = "";
		String fnbr = "";

		int fmt = 0, al = 0, pl = 0, cur = 0;// Formats fmt, al = Alignment, pl
												// = places, cur = Currency
		// Extracting the formats
		if (c.flag != 0) {
			fmt = c.format;
			al = fmt / 100; // Extracting the Alignment
			pl = (fmt % 100) / 10;
			cur = (fmt % 10);
			if (cur == 1) {
				fnbr = "BRL ";
			}
			if (cur == 2) {
				fnbr = "USD ";
			}
			if (cur == 3) {
				fnbr = "EUR ";
			}
			if (cur == 4) {
				fnbr = "YEN ";
			}
			if (cur == 9) {
				fnbr = "%";
			}
		}

		if (c.flag == 0) { // cell empty
			str = "";
		} else if (c.flag == 1) { // Number
			float nb = c.value;
			if (cur == 9) {
				nb = nb * 100;
			} // %

            
            str = nb.ToString("n"+pl.ToString().Replace(" ","")); //Decimal Places
			if (cur >= 1 && cur <= 4) {str = fnbr + str;} // Currency
			if (cur == 9) {	str = str + fnbr;} // %

		} else if (c.flag == 2) { // Text
			str = c.formula;
		}

		int i = 0, len = str.Length;

		int pos = 0;// Default Align Left
		for (i = 0; i < 50; i++)
			prt[i] = ' ';

		// default Align Left;
		pos = 0;
        if (len > size) len = size;;
		// --------- Format Alignment -----------------
		if (al == 1) {pos = 0;} // Align-Left
		if (al == 2) {pos = size - len;} // Align-Right
		if (al == 3) {pos = size / 2 - len / 2;} // Align-Center
		// ----------------------------------------------
		char []strx = str.ToCharArray();
		for (i = 0; i < len; i++) {
			prt[i + pos] = strx[i];
		}
		prt[size] = (char) 0;
		String str2 = "";
		try {
			for (i = 0; i < size; i++) {
				if (prt[i] == 0)
					break;
				str2 = str2 + prt[i].ToString();
			}
		} catch (Exception ex) {
		}
		return str2;

	}

        String printRow(int row)
        {
            String buffer = row.ToString();
            return printData(buffer, 4, 0);

        }

        String printCol(int col)
        {
            return printData(getCol(col), col_size[col], 2);// center
        }

        String printData(String strx, int size, int align)
        {// Commom code
            if (size == 0)
            {
                return "";
            }
            char[] prt = new char[50];
            char[] str = strx.ToCharArray();

            int i = 0, len = str.Length;
            int pos = 0;// Default Align Left
            for (i = 0; i < 50; i++)
                prt[i] = ' ';
            // default Align Left;
            if (align == 0){pos = 0;} // Align-Left
            if (align == 1){pos = size - len;} // Align-Right
            if (align == 2){pos = size / 2 - len / 2;} // Align-Center

            for (i = 0; i < len; i++)
            {
                prt[i + pos] = str[i];
            }
            String str2 = "";
            for (i = 0; i < size; i++)
            {
                if (prt[i] == 0)
                    break;
                str2 = str2 + prt[i].ToString();

            }

            prt[size] = (char) 0;
            return str2;

        }

        
        // ********************** Edit Operations - Insert/Delete/Exclude/Copy Range
        // ******************
        // ************** Insert, Exclude, Delete Delete Rows and Columns
        // ****************************
        void excludeRows(int r1, int r2)
        {
            int r, c, lc, lr;
            lc = lastCol;
            lr = lastRow;

            for (r = r1; r <= lr + 1; r++)
            {
                for (c = 0; c <= lc + 1; c++)
                {
                    copyCell(r, c, r2 + 1, c);
                }
                r2++;
            }
        }

        void excludeColumns(int c1, int c2)
        {
            int r, c, lc, lr;
            lc = lastCol;
            lr = lastRow;
            for (c = c1; c <= lc + 1; c++)
            {
                for (r = 0; r <= lr + 1; r++)
                {
                    copyCell(r, c, r, c2 + 1);
                }
                c2++;
            }
        }

        void deleteRows(int r1, int r2)
        {
            int r, c, lr, lc;
            lc = lastCol;
            lr = lastRow;

            for (r = r1; r <= r2; r++)
            {
                for (c = 0; c <= lc + 1; c++)
                {
                    delCell(r, c);
                }
            }

        }

        void deleteColumns(int c1, int c2)
        {
            int r, c, lr, lc;
            lc = lastCol;
            lr = lastRow;
            for (c = c1; c <= c2; c++)
            {
                for (r = 0; r <= lr + 1; r++)
                {
                    delCell(r, c);
                }
            }
        }

        void insertRow(int r1)
        {
            int r, c, lr, lc;
            lc = lastCol;
            lr = lastRow;
            for (r = lr + 1; r >= r1; r--)
            {
                for (c = 0; c <= lc + 1; c++)
                {
                    copyCell(r, c, r - 1, c);
                }
            }
            for (c = 0; c <= lc + 1; c++)
            {// Cleaning the Current Row;
                delCell(r1, c);
            }

        }

        void insertCol(int c1)
        {
            int r, c, lr, lc;
            lc = lastCol;
            lr = lastRow;
            for (c = lc + 1; c >= c1; c--)
            {
                for (r = 1; r <= lr + 1; r++)
                {
                    copyCell(r, c, r, c - 1);
                }
            }
            for (r = 1; r <= lr + 1; r++)
            {// Cleaning the Current Column;
                delCell(r, c1);
            }

        }

        void copyRange(int r1, int c1, int r2, int c2, int r3, int c3)
        {
            int r, c, dr, dc;
            dr = r3 - r1;
            dc = c3 - c1;

            for (r = r1; r <= r2; r++)
            {
                for (c = c1; c <= c2; c++)
                {
                    copyCell(r + dr, c + dc, r, c);
                }
            }
        }

        void delRange(int r1, int c1, int r2, int c2)
        {
            int r, c;
            for (r = r1; r <= r2; r++)
            {
                for (c = c1; c <= c2; c++)
                {
                    delCell(r, c);
                }
            }
        }

        void HideRC(String str1, String str2, String s1, String s2)
        {
            // Function to Hide/Unhide, columns or Rows
            int flag = 0, n1 = 0, n2 = 0, i = 0;
            if (str1.Equals("H")) { flag = 1; } else { flag = 0; }

            if (str2.Equals("R"))
            {//row
                n1 = Convert.ToInt32(s1);
                n2 = Convert.ToInt32(s2);
                for (i = n1; i <= n2; i++)
                {
                    row_hide[i] = flag;
                }
            }
            else if (str2.Equals("C"))
            {//column
                RC t1 = getRC(s1); if (t1.c == 0) { n1 = Convert.ToInt32(s1); } else { n1 = t1.c; }
                RC t2 = getRC(s2); if (t2.c == 0) { n2 = Convert.ToInt32(s2); } else { n2 = t2.c; }
                for (i = n1; i <= n2; i++)
                {
                    col_hide[i] = flag;
                }

            }


        }
        
        // ***********************************[MENUS]******************************************************
        void help()
        {
            String msg = "";
            msg = msg + "\n" + (" ========= HELP MENU============");
            msg = msg + "\n" + (" [H] - Help Menu                ");
            msg = msg + "\n" + (" [L] - List of Functions        ");
            msg = msg + "\n" + (" [F] - Menu Files Options       ");
            msg = msg + "\n" + (" [C] - Edit Cell's Contents     ");
            msg = msg + "\n" + (" [V] - View Cell's Properties   ");
            msg = msg + "\n" + (" [E] - Menu Edit Options        ");
            msg = msg + "\n" + (" [S] - Edit Columns Size        ");
            msg = msg + "\n" + (" [U] - Hide/Unhide Rows/Cols    ");
            msg = msg + "\n" + (" [T] - Format Range             ");
            msg = msg + "\n" + (" [X] - Test RC Function         ");
            msg = msg + "\n" + (" ================================");
            msgbox(msg, "Basic VisiCalc");

        }

        void formatRange()
        {
            int op = 0;
            String txt1 = "", txt2 = "";
            int r, c, r1, c1, r2, c2, fmt;

            String msg = "";
            msg = msg + "\n" + ("  ==== FORMAT OPTIONS ==== ");
            msg = msg + "\n" + ("  [1XX] - Align Left     ");
            msg = msg + "\n" + ("  [2XX] - Align Right    ");
            msg = msg + "\n" + ("  [3XX] - Align Center   ");
            msg = msg + "\n" + ("  [X0X] - 0 Dec. Places  ");
            msg = msg + "\n" + ("  [X2X] - 2 Dec. Places  ");
            msg = msg + "\n" + ("  [X9X] - 9 Dec. Places  ");
            msg = msg + "\n" + ("  [XX0] - Number Format  ");
            msg = msg + "\n" + ("  [XX1] - Currency BRL   ");
            msg = msg + "\n" + ("  [XX2] - Currency USD   ");
            msg = msg + "\n" + ("  [XX3] - Currency EUR   ");
            msg = msg + "\n" + ("  [XX4] - Currency YEN   ");
            msg = msg + "\n" + ("  [XX9] - Percent Format ");
            msg = msg + "\n" + ("  ========================");
            msg = msg + "\n\n" + ("Inform the Range [CR1 CR2] and format [XXX]: ");
            String[] pr = inputBox("Basic VisiCalc", msg).Split(new string[] { " " }, StringSplitOptions.None);
            if (pr[0].Length == 0) return;

            RC t1 = getRC(pr[0]);
            r1 = t1.r;
            c1 = t1.c;
            RC t2 = getRC(pr[1]);
            r2 = t2.r;
            c2 = t2.c;
            fmt = Convert.ToInt32(pr[2]);

            for (r = r1; r <= r2; r++)
            {
                for (c = c1; c <= c2; c++)
                {
                    tb[r, c].format = fmt;
                }
            }

        }

        void listFunctions()
        {
            int op;
            String msg = "";
            msg = msg + "\n" + ("  ===== LIST OF FUNCTIONS ====   ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + ("  =SUM(RG1:RG2)                  ");
            msg = msg + "\n" + ("      Sums the Range (RG1:RG2)   ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + ("  =SUMIF(CR1:CR2; CDX; RG1:RG2)  ");
            msg = msg + "\n" + ("    Sums the Range (RG1:RG2)     ");
            msg = msg + "\n" + ("    If CDX is true in (CD1:CD2)  ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + ("  =IF(CL1=VL1;CL2;CL3)           ");
            msg = msg + "\n" + ("   IF(CD=1) THEN ELSE            ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + ("  =A1+B2/C3*C4-C5 - Basic Math   ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + ("  =VLOOKUP(VL1;RG1:RG2;DSL)      ");
            msg = msg + "\n" + ("  =HLOOKUP(VL1;RG1:RG2;DSL)      ");
            msg = msg + "\n" + ("   LOOKUP in Columns or Rows     ");
            msg = msg + "\n" + ("                                 ");
            msg = msg + "\n" + (" ================================= ");
            msgbox(msg, "Basic VisiCalc");

        }

        private int msgbox(String msg, String title)
        {
            // JOptionPane.showMessageDialog(null, msg);
            int n = 0;// JOptionPane.showOptionDialog(null, msg, title, JOptionPane.OK_OPTION, JOptionPane.QUESTION_MESSAGE,
                     //null, null, null);

            MessageBox.Show(msg, title);
            return n;
        }

        // *********************[CELL FUNCTIONS - CALCULATE, DECODE AND COPY
        // FUNCTIONS]********************
        static void calculateCells()
        {
            int r, c, lr, lc;
            float vl;
            lr = lastRow;
            lc = lastCol;

            for (r = 0; r <= lr; r++)
            {
                for (c = 0; c <= lc; c++)
                {
                    try
                    {
                        if (tb[r, c].flag != 0)
                        {
                            if (tb[r, c].formula.Length > 0)
                            {
                                vl = cellFunctions(r, c);
                                // System.out.println("cell ("+r+","+c+") = "+
                                // tb[r, c].value);
                                tb[r, c].value = vl;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        //System.out.println("ERROR[CalculateCells]: cell (" + getCol(c) + "" + r + ") - " + ex.getMessage());

                    }
                }
            }
        }

        static float cellFunctions(int r, int c) {
		// System.out.println("Cell Functionss: (R="+r+"C="+c+")"+
		// tb[r, c].formula);

		char[] tmp = new char[50];
		int i, j, len, i1 = 0, i2 = 0;
		float tf = 0;
		String[] strx = new String[20];
		for (i = 0; i < 20; i++)
			strx[i] = "";

		String tmpx2 = tb[r, c].formula;
		char [] tmp2 = tmpx2.ToCharArray();
		len = tmpx2.Length;

		int ix = 0;
		for (i = 0; i < len; i++) { // REMOVING SPACES / ToUpper
			if (tmp2[i] != ' ') {
				tmp[ix] = Char.ToUpper(tmp2[i]);
				ix++;
			}
		}

		tmp[len] = (char) 0;
		len = ix;
		strx[0] = "";
		i2 = 0;
		for (i = 0; i < len; i++) {// Formating the Formula
			if (tmp[i] == 0) {
				break;
			}
			if (tmp[i] == '=' || tmp[i] == '+' || tmp[i] == '-' || tmp[i] == '/' || tmp[i] == '*' || tmp[i] == '!'
					|| tmp[i] == '>' || tmp[i] == '<') {
				if (i2 > 0)
					i1++;
				strx[i1] = tmp[i].ToString();
				i1++;
				strx[i1] = "";
				i2 = 0;
			} else if (tmp[i] == ' ' || tmp[i] == '\n' || tmp[i] == '(' || tmp[i] == ')' || tmp[i] == ':'
					|| tmp[i] == ';') {
				i1++;
				strx[i1] = "";
				i2 = 0;
			} else {
				strx[i1] = strx[i1] + tmp[i].ToString();
				i2++;
			}
		}
		if (strx[0].Equals("=")) {
		} else {
			return 0;
		}

		// ----------- Checking ---------------
		int ft = 0; // Flag Test
		if (ft == 1) {
			if (strx[0].Equals("=")) {
				//System.out.printf("\nDiv: ");
				for (i = 0; i <= i1; i++) {
					//System.out.printf("'%s' /", strx[i]);
				}
				//System.out.println("");
			}
		}
		// -------------------------------------

		// =SUM (R1C1 R2C2)
		if (strx[1].Equals("SUM")) {
			// printf("\nFormula: ",tmp);

			RC t1 = getRC(strx[2]);
			RC t2 = getRC(strx[3]);
			// Sum Range
			tf = 0;
			for (r = t1.r; r <= t2.r; r++) {
				for (c = t1.c; c <= t2.c; c++) {
					tf = tf + tb[r, c].value;
				}
			}
			return tf;
		}
		// =SUMIF (CD1:CD2;CDX;RG1:RG2)
		if (strx[1].Equals("SUMIF")) {
			float a1;
			int d1, d2;
			// = SUM(CD1:CD2;CDX;RG1:RG2)
			// 0 1 2 3 4 5 6

			RC t1 = getRC(strx[2]);
			RC t2 = getRC(strx[3]);
			float cdx = getCellVal(strx[4]);
			RC t3 = getRC(strx[5]);
			RC t4 = getRC(strx[6]);

			d1 = t3.r - t1.r;
			d2 = t3.c - t1.c;

			// Sum Range
			tf = 0;
			for (r = t1.r; r <= t2.r; r++) {
				for (c = t1.c; c <= t2.c; c++) {
					a1 = tb[r, c].value;
					if (a1 == cdx) {
						tf = tf + tb[r + d1, c + d2].value;
					}
				}
			}
			return tf;
		}

		// =VLOOKUP(VS;X1:X2;C)
		if (strx[1].Equals("VLOOKUP")) {
			float a1, a2;
			int dsl = 0;
			// =VLOOKUP(VS;X1:X2;C)
			// 0 1 2 3 4 5
			tf = 0;
			a1 = getCellVal(strx[2]);// Value Searched
			RC t1 = getRC(strx[3]); // Range X1
			RC t2 = getRC(strx[4]); // Range X2
			dsl = (int) getCellVal(strx[5]);// DSL Column

			// Sum Range
			tf = 0;
			for (r = t1.r; r <= t2.r; r++) {
				a2 = tb[r, t1.c].value;
				if (a1 == a2) {
					tf = tb[r, t1.c + dsl - 1].value;
					break;
				}
			}
			return tf;
		}
		// =HLOOKUP(VS;X1:X2;C)
		if (strx[1].Equals("HLOOKUP")) {
			float a1, a2;
			int dsl = 0;
			// =HLOOKUP(VS;X1:X2;C)
			// 0 1 2 3 4 5
			tf = 0;
			a1 = getCellVal(strx[2]);// Value Searched
			RC t1 = getRC(strx[3]); // Range X1
			RC t2 = getRC(strx[4]); // Range X2
			dsl = (int) getCellVal(strx[5]);// DSL Column

			tf = 0;
			for (c = t1.c; c <= t2.c; c++) {
				a2 = tb[t1.r, c].value;
				if (a1 == a2) {
					tf = tb[t1.r + dsl - 1, c].value;
					break;
				}
			}
			return tf;
		}
		if (strx[1].Equals("IF")) {
			// =IF(A = B):THEN:ELSE
			// 0 1 2 3 4 5 6
			float a, b;
			int flag = 0;

			a = getCellVal(strx[2]);
			b = getCellVal(strx[4]);

			if (strx[3].Equals("=")) {
				if (a == b) {
					flag = 1;
				}
			}
			;
			if (strx[3].Equals("!")) {
				if (a != b) {
					flag = 1;
				}
			}
			;
			if (strx[3].Equals(">")) {
				if (a > b) {
					flag = 1;
				}
			}
			;
			if (strx[3].Equals("<")) {
				if (a < b) {
					flag = 1;
				}
			}
			;

			if (flag == 1) {
				tf = getCellVal(strx[5]);
			} else {
				tf = getCellVal(strx[6]);
			}
			return tf;
		}

		// = R1C1 + R2C2 - R3C2 - R4C4 * R5C5 / R6C6
		if (strx[0].Equals("=")) {
			int op = 1; // 0 = Equal 1 Addition, 2 - Subtraction , 3 =
						// Multiplication 4 = Division
			float val = 0;
			tf = 0;

			for (i = 1; i <= i1 + 1; i++) {
				try {
					if (strx[i].Length == 1) {
						if (strx[i].Equals("="))
							op = 0;
						if (strx[i].Equals("+"))
							op = 1;
						if (strx[i].Equals("-"))
							op = 2;
						if (strx[i].Equals("*"))
							op = 3;
						if (strx[i].Equals("/"))
							op = 4;
					}
					if (strx[i].Length > 1) {
						try {
							val = getCellVal(strx[i]);

							if (op == 0) {
								tf = val;
							}
							if (op == 1) {
								tf = tf + val;
							}
							if (op == 2) {
								tf = tf - val;
							}
							if (op == 3) {
								tf = tf * val;
							}
							if (op == 4) {
								tf = tf / val;
							}
						} catch (Exception ex) {
							val = 0;
						}
						// System.out.printf("P= %.2f / T= %.2f \n", val, tf);
					}
				} catch (Exception ex) {
				}

			}
			return tf;
		}
		return 0;
	}

        void copyFormula(int rd, int cd, int ro, int co) {
		char[] tmp = new char[50];
		int i, j, len, i1 = 0, i2 = 0;
		float tf = 0;
		String[] strx = new String[20];
		for (i = 0; i < 20; i++)
			strx[i] = "";

		String tmpx2 = tb[ro, co].formula;
		char[] tmp2 = tmpx2.ToCharArray();
		//
		len = tmpx2.Length;
		int ix = 0;
		for (i = 0; i < len; i++) { // REMOVING SPACES / ToUpper
			if (tmp2[i] != ' ') {
				tmp[ix] = Char.ToUpper(tmp2[i]);
				ix++;
			}
		}
		tmp[len] = (char) 0;

		if (tmp[0] != '=') {
			tb[rd, cd].formula = tb[ro, co].formula;
		}

		// System.out.println("Formula = " + String.valueOf(tmp));

		len = ix;
		strx[0] = "";
		for (i = 0; i < len; i++) {// Formating the Formula
			if (tmp[i] == 0) {
				break;
			}
			if (tmp[i] == '=' || tmp[i] == '+' || tmp[i] == '-' || tmp[i] == '/' || tmp[i] == '*' || tmp[i] == '!'
					|| tmp[i] == '>' || tmp[i] == '<' || tmp[i] == ' ' || tmp[i] == '\n' || tmp[i] == '('
					|| tmp[i] == ')' || tmp[i] == ':' || tmp[i] == ';') {
				if (i2 > 0)
					i1++;
				strx[i1] = tmp[i].ToString();
				i1++;
				strx[i1] = "";
				i2 = 0;
			} else {
				strx[i1] = strx[i1] + tmp[i].ToString();
				i2++;
			}
		}

		// ----------- Checking ---------------
		int ft = 0; // Flag Test
		if (ft == 1) {
			// if (strx[0].Equals("=")) {
			//System.out.printf("\nDiv: ");
			for (i = 0; i < i1; i++) {
				//System.out.printf("'%s' /", strx[i]);
			}
			//System.out.println("");
			// }
		}
		// -------------------------------------

		// = R1C1 + R2C2 - R3C2 - R4C4 * R5C5 / R6C6
		int dr, dc, r, c;
		if (strx[0].Equals("=")) {
			try {
				for (i = 1; i <= i1; i++) {
					if (strx[i].Length > 1) {
						char[] tc = strx[i].ToCharArray();
						if (tc[0] >= 48 && tc[0] <= 57 || tc[0] == '[') {
							// Literal - Keep it as it is
						} else {
							RC t1 = getRC(strx[i]);// Cell
							if (t1.r != 0 && t1.c != 0) {
								// Change Reference
								dr = t1.r - ro; // Ref - Orige
								dc = t1.c - co; // Ref - Orige
								r = rd + dr;
								c = cd + dc;
								strx[i] = getCell(r, c);
							}
						}
					}
				}
			} catch (Exception ex) {
			}

			String Tmpx = "";
			for (i = 0; i <= i1 + 1; i++) {
				Tmpx = Tmpx + strx[i];
			}

			// Copying the Changed String to the New Cell
			// System.out.println("New Formula = " + Tmpx);
			tb[rd, cd].formula = Tmpx;
		} else {
			tb[rd, cd].formula = tb[ro, co].formula;
		}

	}

        // **************** [FILES: OPEN, SAVE, SAVE AS, CLOSE] **************** 

        
        void files(int op)
        {
            if (op == 1)
            {
                newSpreadsheet();
            }
            if (op == 2)
            {
                openFile();
            }
            if (op == 3)
            {
                saveFile();
            }
            if (op == 4)
            {
                saveFileAs();
            }
            if (op == 5)
            {
                this.Close();
                this.Dispose();
            }
        }

        
        void saveFile() {
		String cFile = currPath + currFile;

		// lastCol = NCOL;
		// lastRow = NROW;

		
		try {
            StreamWriter file1 = new StreamWriter(cFile);
			int r, c;

			file1.Write("VisicalcProjectFile:\r\n");
			file1.Write("RC:\r\n " + lastRow + 1 + " " + lastCol + 1 + " \r\n");
			for (c = 0; c < lastCol + 1; c++) {
				file1.Write("col:\r\n " + c + " " + col_size[c] + " " + col_hide[c] + " \r\n");
			}

			for (r = 0; r < lastRow + 1; r++) {
				file1.Write("row:\r\n " + r + " " + row_hide[r] + " \r\n");
			}

			int nr = 0;
			for (r = 0; r <= lastRow + 1; r++) {
				for (c = 0; c <= lastCol + 1; c++) {
					if (tb[r, c].flag != 0) {
						nr++;
					}
				}
			}

			file1.Write("nr= " + nr + "\r\n");// number of cells

			for (r = 0; r <= lastRow + 1; r++) {
				for (c = 0; c <= lastCol + 1; c++) {
					if (tb[r, c].flag != 0) {
						file1.Write("data1:\r\n " + r + " " + c + " " + tb[r, c].flag + " " + tb[r, c].format + " "
								+ tb[r, c].value + " \r\n");
						if (tb[r, c].formula.Length > 0) {
							file1.Write("data2:\r\n" + tb[r, c].formula + "\r\n");
						}
					}
				}
			}
			file1.Write("End:\r\n");
			file1.Close();

			String msg = "";
			msg = msg + "\n" + (" ********************** ");
			msg = msg + "\n" + (" ** ==File SAVED! == ** ");
			msg = msg + "\n" + (" ********************** ");
			msgbox(msg, "Basic VisiCalc");
		} catch (IOException e) {
			msgbox("Writing to File Error: \n" + e.Message, "Error Saving File");
		}

	}

        void saveFileAs()
        {
            String tmp = inputBox("Save File As", "\n Inform the new name for the file: [" + currFile + "]:");
            if (tmp.Length > 0)
            {
                currFile = tmp;
            }
            saveFile();

        }

        void openFile() {
		String tmp = inputBox("Open File", "Open the Current File [" + currFile + "] ? [Y/N]");
		if (tmp.Equals("N") || tmp.Equals("n")) {
			tmp = inputBox("Open File", "\nInform the Name of the File: ");
			if (tmp.Length > 0) {
				currFile = tmp;
			}
		}
		newSpreadsheet();
		// ****************** Opening the File ******************
		String cFile = "";
		String tmpf = "";
		cFile = currPath + currFile;
		int r = 0, c = 0, a = 0, b = 0, ch = 0;
		float f;
		// printf("\n%s\n", cFile);

		try {

			StreamReader file2 = new StreamReader(cFile);

			while ((tmpf = file2.ReadLine()) != null) {

				if (tmpf.Equals("RC:")) {
					String [] pr = file2.ReadLine().Split(new string[] { " " }, StringSplitOptions.None);
					lastRow = 0;//Convert.ToInt32(pr[1]);
					lastCol = 0;//Convert.ToInt32(pr[2]);

				}
				if (tmpf.Equals("col:")) {
					String [] pr = file2.ReadLine().Split(new string[] { " " }, StringSplitOptions.None);
					a = Convert.ToInt32(pr[1]);
					b = Convert.ToInt32(pr[2]);
					c = Convert.ToInt32(pr[3]);
					col_size[a] = b;
					col_hide[a] = c;

				}
				if (tmpf.Equals("row:")) {
					String [] pr = file2.ReadLine().Split(new string[] { " " }, StringSplitOptions.None);
					a = Convert.ToInt32(pr[1]);
					b = Convert.ToInt32(pr[2]);
					row_hide[a] = b;
				}
				if (tmpf.Equals("data1:")) {
					// fscanf(file2," %i %i %i %i %f \n", &r, &c, &a, &ch, &f);
					String [] pr= file2.ReadLine().Split(new string[] { " " }, StringSplitOptions.None);
					r = Convert.ToInt32(pr[1]);
					c = Convert.ToInt32(pr[2]);
					a = Convert.ToInt32(pr[3]);
					ch = Convert.ToInt32(pr[4]);
					f = float.Parse(pr[5].Replace(".",","));

					if (r != 0 && c != 0) {
						tb[r, c].flag = a;
						tb[r, c].format = ch;
						tb[r, c].value = f;
						if (r>lastRow) lastRow = r;
						if (c>lastCol) lastCol = c;
					}
				}
				if (tmpf.Equals("data2:")) {
					String str = file2.ReadLine();
					if (r != 0 && c != 0) {
						if (str.Length > 0) {
							tb[r, c].formula = str;
						}
					}
				}
				if (tmpf.Equals("End:")) {
					//System.out.println("\n**End of File**\n");
                    showMatrix();
                    msgbox("File " + currFile + " Loaded!", "File Open");
					break;
				}
			}
			file2.Close();

		} catch (IOException e) {
			msgbox("Error: \n" + e.Message, "File Open Error");
		}

	}
        //*/
        // ************************************************************************************************

        public void updateCell()
        {

            // String str = inputBox("Basic VisiCalc", "Edit Cells \n\n [Cell]
            // [Value] [Format/Text]");
            String txt1 = t1.Text;
            String txt2 = t2.Text;
            String txt3 = t3.Text;

            RC tx = new RC(0, 0);
            tx = getRC(txt1);

            if (tx.r > 0 && tx.c > 0)
            {// Formula/Text
                if (txt3.Length > 0)
                {
                    String x = txt3.Substring(0, 1);
                    if (x.Equals("="))
                    {
                        tb[tx.r, tx.c] = setCell(1, 0, 0, txt3);
                    }
                    else
                    {
                        tb[tx.r,tx.c] = setCell(2, 0, 0, txt3);
                    }
                    // System.out.println("Cell: "+ txt1 + ", (r= " + tx.r + ", c=
                    // "+ tx.c+") Text = "+ tb[tx.r, tx.c].formula);
                    // System.out.println("(r= " + tx.r + ", c= "+ tx.c+") getCell =
                    // " + getCell(tx.r, tx.c) + ", getCellValue = " +
                    // getCellVal(txt1));
                }
                else
                {// Value
                    float f = float.Parse(txt2);
                    tb[tx.r, tx.c] = setCell(1, 0, f, "");
                    // System.out.println("Cell: "+ txt1 + ", (r= " + tx.r + ", c=
                    // "+ tx.c+") Value = "+ tb[tx.r, tx.c].value);
                    // System.out.println("(r= " + tx.r + ", c= "+ tx.c+") getCell =
                    // " + getCell(tx.r, tx.c) + ", getCellValue = " +
                    // getCellVal(txt1));
                }
                if (tx.r > lastRow)
                    lastRow = tx.r;
                if (tx.c > lastCol)
                    lastCol = tx.c;
                t2.Text = "";
                t3.Text = "";
                showMatrix();

            }
        }

        void editOptions(int op)
        {
            int r1, c1, r2, c2, r3, c3;
            String txt1 = "";
            String txt2 = "";
            String txt3 = "";
            String msg = "";

            if (op == 0)
            {

                msg = msg + "\n" + (" =========== EDIT OPTIONS ===========");
                msg = msg + "\n" + (" 1 - Insert Row   \t2 - Insert Col   ");
                msg = msg + "\n" + (" 3 - Delete Rows  \t4 - Delete Cols  ");
                msg = msg + "\n" + (" 5 - Excl.  Rows  \t6 - Excl. Cols   ");
                msg = msg + "\n" + (" 7 - Delete Range \t8 - Copy Range   ");
                msg = msg + "\n" + (" 9 - Format Range \t10- Resize Cols  ");
                msg = msg + "\n" + (" 11- Edit Cells   \t12- (Un)Hide RCs ");
                msg = msg + "\n" + (" 13- View Cells   \t                 ");
                msg = msg + "\n" + (" ====================================");

                msg = msg + "\n\n" + ("Enter an Option: ");
                op = 0;
                try
                {
                    op = Convert.ToInt32(inputBox("Edit Options", msg));
                }
                catch (Exception ex) { op = 0; return;}

                // lastCol = NCOL-2;
                // lastRow = NROW-2;

            }

            if (op == 1)
            {
                r1 = Convert.ToInt32(inputBox("Insert Row", "\nEnter the Row number:"));
                insertRow(r1);
            }
            if (op == 2)
            {
                txt1 = (inputBox("Insert Col", "\nEnter the Column:"));
                RC t = getRC(txt1);
                c1 = t.c;
                insertCol(c1);
            }
            if (op == 3)
            {
                String[] pr = (inputBox("Delete Rows [R1 R2]", "\nEnter [r1] [r2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                r1 = Convert.ToInt32(pr[0]);
                r2 = Convert.ToInt32(pr[1]);

                deleteRows(r1, r2);
            }
            if (op == 4)
            {
                String[] pr = (inputBox("Delete Cols [C1 C2]", "\nEnter [C1] [C2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                RC t1 = getRC(pr[0]);
                c1 = t1.c;
                RC t2 = getRC(pr[1]);
                c2 = t2.c;
                deleteColumns(c1, c2);
            }

            if (op == 5)
            {
                String[] pr = (inputBox("Exclude Rows [R1 R2]", "\nEnter [r1] [r2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                r1 = Convert.ToInt32(pr[0]);
                r2 = Convert.ToInt32(pr[1]);
                if (r1 != 0 && r2 != 0)
                {
                    excludeRows(r1, r2);
                }
            }
            if (op == 6)
            {
                String[] pr = (inputBox("Exclude Cols [C1 C2]", "\nEnter [C1] [C2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                RC t1 = getRC(pr[0]);
                c1 = t1.c;
                RC t2 = getRC(pr[1]);
                c2 = t2.c;
                if (c1 != 0 && c2 != 0)
                {
                    excludeColumns(c1, c2);
                }
            }

            if (op == 7)
            {
                String[] pr = (inputBox("Delete Range[RC1 RC2]", "\nEnter [RC1] [RC2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                RC t1 = getRC(pr[0]);
                r1 = t1.r;
                c1 = t1.c;
                RC t2 = getRC(pr[1]);
                r2 = t2.r;
                c2 = t2.c;
                delRange(r1, c1, r2, c2);
            }
            if (op == 8)
            {
                String[] pr = (inputBox("Copy Range (RC1 RC2] to [RC3]", "\nEnter [RC1] [RC2] [RC3]:")).Split(new string[] { " " }, StringSplitOptions.None);

                RC t1 = getRC(pr[0]);
                r1 = t1.r;
                c1 = t1.c;
                RC t2 = getRC(pr[1]);
                r2 = t2.r;
                c2 = t2.c;
                RC t3 = getRC(pr[2]);
                r3 = t3.r;
                c3 = t3.c;

                copyRange(r1, c1, r2, c2, r3, c3);
            }

            if (op == 9)
            {
                formatRange();
            }
            if (op == 10)
            {
                otherOptions('S');
            }
            if (op == 11)
            {
                otherOptions('C');
            }
            if (op == 12)
            {
                otherOptions('U');
            }
            if (op == 13)
            {
                otherOptions('V');
            }

            showMatrix();
        }

        void otherOptions(char kx)
        {
            // Format Columns, Resize Columns, Edit Cells
            String txt1 = "";
            String txt2 = "";
            String msg = "";

            if (kx == 'U' || kx == 'u')
            {// Hide and Unhide
                String[] pr = (inputBox("Hide/UnHide RC", "\nEnter [H/U] [R/C] [N1 N2]:")).Split(new string[] { " " }, StringSplitOptions.None);
                if (pr[0].Length > 0) { HideRC(pr[0].ToUpper(), pr[1].ToUpper(), pr[2], pr[3]); }

            }
            if (kx == 'S' || kx == 's')
            {
                int c, c1, c2, sz;
                String[] pr = (inputBox("Coluns Size [C1..C2] ", "\nEnter [C1] [C2] [size]:")).Split(new string[] { " " }, StringSplitOptions.None);

                RC t1 = getRC(pr[0]);
                c1 = t1.c;
                RC t2 = getRC(pr[1]);
                c2 = t2.c;
                sz = Convert.ToInt32(pr[2]);
                for (c = c1; c <= c2; c++)
                {
                    col_size[c] = sz;
                }
            }
            if (kx == 'C' || kx == 'c')
            {
                updateCell();
            }
            if (kx == 'V' || kx == 'v')
            {// View Cell's Properties

                txt1 = inputBox("View Cell", "\nEnter [Cell]:");

                while (txt1.Length > 0)
                {

                    msg = "";
                    RC t = getRC(txt1);
                    msg = msg + "\nCell   : " + txt1.ToUpper() + " [R=" + t.r + ", C=" + t.c + "]";
                    msg = msg + "\nFlag   : " + tb[t.r, t.c].flag;
                    msg = msg + "\nFormat : " + tb[t.r, t.c].format;
                    msg = msg + "\nValue  : " + tb[t.r, t.c].value;
                    msg = msg + "\nFormula: " + tb[t.r, t.c].formula;
                    msg = msg + "\n\nEnter [Cell]:";

                    txt1 = inputBox("View Cell", msg);
                }
            }

        }

        static String inputBox(String title, String prompt)
        {
            String str = "";
            // str = (JOptionPane.showInputDialog(prompt));
            //str = JOptionPane.showInputDialog(null, prompt, title, JOptionPane.PLAIN_MESSAGE);
            try { str = InputBOX(title, prompt); } catch (Exception ex) { }
            return str;
        }

        //*/

        public static String InputBOX(string title, string promptText)
        {
            //code from http://www.csharp-examples.net/inputbox/ - 24/06/2016 23:00
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = "";

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            String[] pr = promptText.Split(new string[] { "\n" }, StringSplitOptions.None);
            int lb_h = (pr.Count()+1)*13;

            label.SetBounds(10, 10, 250, lb_h);
            int lh = label.Height;
            label.AutoSize = true;


            int dsl = 10 + label.Height;
            textBox.SetBounds(10, 5+dsl, 250, 20);
            buttonOk.SetBounds(50, 30+dsl, 60, 23);
            buttonCancel.SetBounds(125, 30+dsl, 60, 23);

            //label.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            //textBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            //buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            //buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(label.Width+20, dsl + 60);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.AutoSize = true;

            //form.ClientSize = new Size(Math.Max(320, label.Width + 20), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;


            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;


            DialogResult dialogResult = form.ShowDialog();
            String str = textBox.Text;
            return str;


        }

        //********************* Buttons *********************************

        private void btnSave_Click(object sender, EventArgs e)
        {
            files(3);
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            newSpreadsheet();
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            help();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            editOptions(0);
        }
        private void menuFile_Click(object sender, EventArgs e)
        {
            String sel = sender.ToString();
            switch (sel)
            {
                case "&New File": files(1); break;
                case "&Open File": files(2); break;
                case "&Save File": files(3); break;
                case "Save File &As": files(4); break;

                default:
                    msgbox("MenuFile\n Sender: " + sender.ToString() + "\ne = " + e.ToString(), "Basic VisiCalc");
                    break;
            }

        }
        private void menuEdit_Click(object sender, EventArgs e)
        {
            String sel = sender.ToString();
            switch (sel)
            {
                case "&Edit Cell": updateCell(); break;
                case "&View Cell": editOptions(13); break;

                case "&Format Range": editOptions(9); break;
                case "&Delete Range": editOptions(7); break;
                case "&Copy Range": editOptions(8); break;
                case "&Resize Columns": editOptions(10); break;
                case "&Hide/Unhide Rows/Cols": editOptions(12); break;

                case "&Insert Row": editOptions(1); break;
                case "I&nsert Column": editOptions(2); break;
                case "De&lete Row(s)": editOptions(3); break;
                case "Dele&te Column(s)": editOptions(4); break;
                case "E&xclude Row(s)": editOptions(5); break;
                case "Excl&ude Column(s)": editOptions(6); break;

                default:
                    msgbox("MenuEdit\n Sender: "+sender.ToString()+"\ne = "+e.ToString(), "Basic VisiCalc");
                    break;
            }


        }
        private void menuHelp_Click(object sender, EventArgs e)
        {
            String sel = sender.ToString();
            switch (sel)
            {
                case "Show &Help": help(); break;
                case "&Functions": listFunctions(); break;
                case "&About":
                    String msg = "";
                        msg += " Basic Visicalc SpreadSheet  \n";
                        msg += " Part of C# Training Program \n";
                        msg += " Made by P.Ramos @ Jun/2016  \n";
                        msgbox(msg, "About");
                    
                    break;
                default:
                    msgbox("MenuHelp\n Sender: " + sender.ToString() + "\ne = " + e.ToString(), "Basic VisiCalc");
                    break;
            }
        }

        private void dbGridTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    } // End of Class
    public class cell
    {
        public int format;
        public String formula;
        public float value;
        public int flag; // 0 - Empty, 1 - Value (Float), 2 - Text

        public cell()
        {// Constructor
            format = 0;
            formula = "";
            value = 0;
            flag = 0;
        }
    }

    public class RC
    {
        public int r;
        public int c;

        public RC(int r1, int c1)
        {
            r = r1;
            c = c1;
        }
    }



}//End of Namespace
