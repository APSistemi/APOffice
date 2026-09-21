$text = Get-Content "c:\ApProject\APOffice\frmGesVariazioni.cs" -Raw

# Fix using declarations
$text = $text.Replace(
"			using Font f = new Font(`"Segoe UI Emoji`", 12f, FontStyle.Regular);
			using SolidBrush b = new SolidBrush(fallbackColor);
			StringFormat fmt = new StringFormat();
			fmt.Alignment = StringAlignment.Center;
			fmt.LineAlignment = StringAlignment.Center;
			g.DrawString(unicodeChar, f, b, new RectangleF(0f, 0f, 24f, 24f), fmt);",
"			using (Font f = new Font(`"Segoe UI Emoji`", 12f, FontStyle.Regular))
            using (SolidBrush b = new SolidBrush(fallbackColor))
            {
			    StringFormat fmt = new StringFormat();
			    fmt.Alignment = StringAlignment.Center;
			    fmt.LineAlignment = StringAlignment.Center;
			    g.DrawString(unicodeChar, f, b, new RectangleF(0f, 0f, 24f, 24f), fmt);
            }"
)

# Fix recursive patterns
$text = $text.Replace(
"if (row.DataBoundItem is DataRowView { Row: not null } drv && drv.Row.RowState != DataRowState.Deleted)",
"DataRowView drv = row.DataBoundItem as DataRowView; if (drv != null && drv.Row != null && drv.Row.RowState != DataRowState.Deleted)"
)

# Fix CurrencyManager pattern
$text = $text.Replace(
"if (dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] is CurrencyManager { Position: >=0 } cm && dgv1.CurrentRow != null && cm.List[dgv1.CurrentRow.Index] is DataRowView)",
"CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager; if (cm != null && cm.Position >= 0 && dgv1.CurrentRow != null && cm.List[dgv1.CurrentRow.Index] is DataRowView)"
)

$text | Set-Content "c:\ApProject\APOffice\frmGesVariazioni.cs"
