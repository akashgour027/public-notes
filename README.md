# public-notes

# default.aspx

    <div class="row">
        <div class="col-md-3">
        	<asp:FileUpload ID="WordFileUpload" runat="server"/>
        </div>
        <div class="col-md-2">
			<asp:Button ID="Upload" runat="server" Text="Upload" OnClick="Upload_Click"/>
        </div>
    </div>
    <div class="row">
		<asp:Label ID="lblError" runat="server" ForeColor="#FF3300"></asp:Label>
    </div>
	<div class="row" style="padding:5px;">
        <div class="col-md-3">
		</div>
	</div>
    <div class="row">
		<asp:GridView ID="GridViewFiles" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataSourceID="SqlDataSource1" GridLines="Horizontal" OnRowCommand="GridViewFiles_RowCommand">
			<AlternatingRowStyle BackColor="#F7F7F7" />
			<Columns>
				<asp:ButtonField CommandName="Download" Text="Download" >
				<ItemStyle Width="100px" />
				</asp:ButtonField>
				<asp:BoundField DataField="Id" HeaderText="Id" InsertVisible="False" ReadOnly="True" SortExpression="Id">
				<ItemStyle Width="80px" />
				</asp:BoundField>
				<asp:BoundField DataField="UploadDate" DataFormatString="{0:dd/MM/yyyy HH:mm}" HeaderText="UploadDate" SortExpression="UploadDate">
				<HeaderStyle Width="100px" />
				<ItemStyle Width="150px" />
				</asp:BoundField>
				<asp:BoundField DataField="FileName" HeaderText="FileName" SortExpression="FileName" />
			</Columns>
			<FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
			<HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
			<PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
			<RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
			<SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
			<SortedAscendingCellStyle BackColor="#F4F4FD" />
			<SortedAscendingHeaderStyle BackColor="#5A4C9D" />
			<SortedDescendingCellStyle BackColor="#D8D8F0" />
			<SortedDescendingHeaderStyle BackColor="#3E3277" />
		</asp:GridView>
		<asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:SunPharmaConnectionString %>" SelectCommand="SELECT [FileName], [UploadDate], [Id] FROM [UploadedFiles]"></asp:SqlDataSource>
    </div>
    
    
    
    
    # default.aspx.cs
    
    
    protected void Upload_Click(object sender, EventArgs e)
		{
			lblError.Text = string.Empty;

			if (WordFileUpload.PostedFile.FileName.Length == 0) 
			{
				lblError.Text = "Please select file to upload.";
				return;
			}

			try
			{
				using (WordprocessingDocument wdoc = WordprocessingDocument.Open(WordFileUpload.PostedFile.InputStream, false))
				{
				}
			}
			catch {
				lblError.Text = "Invalid word file.";
				return;
			}

			Stream fs = WordFileUpload.PostedFile.InputStream;
			BinaryReader br = new BinaryReader(fs);
			byte[] data = br.ReadBytes((int)fs.Length);			

			SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["SunPharmaConnectionString"].ConnectionString);
			con.Open();
			try
			{
				string insertCmd = "INSERT INTO UploadedFiles (FileName, FileData) VALUES(@FileName, @FileData)";
				var com = new SqlCommand(insertCmd, con);
				com.Parameters.Add("@FileName", SqlDbType.VarChar).Value = Path.GetFileName(WordFileUpload.PostedFile.FileName);
				com.Parameters.Add("@FileData", SqlDbType.VarBinary).Value = data;
				com.ExecuteNonQuery();
				GridViewFiles.DataBind();
			}
			finally
			{
				con.Close();
			}
		}

		protected void GridViewFiles_RowCommand(object sender, GridViewCommandEventArgs e)
		{
			string id = GridViewFiles.Rows[Convert.ToInt32(e.CommandArgument)].Cells[0].Text;

			SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["SunPharmaConnectionString"].ConnectionString);
			con.Open();
			try
			{
				string insertCmd = "SELECT FileName, FileData FROM UploadedFiles WHERE Id = @Id";
				var com = new SqlCommand(insertCmd, con);
				com.Parameters.Add("@Id", SqlDbType.Int).Value = id;
				SqlDataReader dataReader = com.ExecuteReader();
				dataReader.Read();
				string fileName = dataReader["FileName"].ToString();
				byte[] fileData = (byte[])dataReader["FileData"];

				using (MemoryStream ms = new MemoryStream())
				{
					ms.Write(fileData, 0, (int)fileData.Length);

					using (WordprocessingDocument wdoc = WordprocessingDocument.Open(ms, true))
					{
						ApplyFooter(wdoc, "By Akash");
					}


					//Respond
					Response.Clear();
					Response.Buffer = true;
					Response.Charset = string.Empty;
					Response.Cache.SetCacheability(HttpCacheability.NoCache);
					Response.ContentType = "";
					Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
					Response.BinaryWrite(ms.ToArray());
					Response.Flush();
					Response.End();
				}
			}
			finally
			{
				con.Close();
			}
		}

		public void ApplyFooter(WordprocessingDocument doc, string footerText)
		{
			// Get the main document part.
			MainDocumentPart mainDocPart = doc.MainDocumentPart;

			FooterPart footerPart1 = mainDocPart.AddNewPart<FooterPart>("r98");


			Footer footer1 = new Footer();
			Paragraph paragraph1 = new Paragraph() { };

			Run run1 = new Run();
			Text text1 = new Text();
			text1.Text = footerText;

			run1.Append(text1);
			paragraph1.Append(run1);
			footer1.Append(paragraph1);
			footerPart1.Footer = footer1;

			SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
			if (sectionProperties1 == null)
			{
				sectionProperties1 = new SectionProperties() { };
				mainDocPart.Document.Body.Append(sectionProperties1);
			}
			FooterReference footerReference1 = new FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };

			sectionProperties1.InsertAt(footerReference1, 0);
		}

    

