using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Reflection;
using System.Windows.Forms;

namespace VisualStudioAddInDemo
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;

            if (
                  connectMode == ext_ConnectMode.ext_cm_AfterStartup
                || connectMode == ext_ConnectMode.ext_cm_UISetup
                || connectMode == ext_ConnectMode.ext_cm_Startup
               )
            {
                try
                {
                    CreateContextMenu(_applicationObject);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
			
		}            

        private void CreateContextMenu(DTE2 _applicationObject)
        {
            try
            {
                Commands2 commands = (Commands2)_applicationObject.Commands;
                CommandBar oCommandBar = ((CommandBars)_applicationObject.CommandBars)["Code Window"];
                AddThisAddin(oCommandBar, _applicationObject);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void AddThisAddin(CommandBar oCommandBar, DTE2 _applicationObject)
        {
            try
            {
                CommandBarButton oControl = (CommandBarButton)
            oCommandBar.Controls.Add(MsoControlType.msoControlButton,
            System.Reflection.Missing.Value,
            System.Reflection.Missing.Value,
            1, true);
                oControl.Caption = "Link Caption";
                oControl.TooltipText = "Insert copied the link as caption";
                oControl.Click += new _CommandBarButtonEvents_ClickEventHandler(oControl_Click);
                link
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
         }

        private void oControl_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            
            try
            {
                TextSelection selection = (TextSelection)_applicationObject.ActiveDocument.Selection;
                TextPoint point = (TextPoint)selection.ActivePoint;
                // Discover every code element containing the insertion point.
                string elems = "";
                vsCMElement scopes = 0;

                foreach (vsCMElement scope in Enum.GetValues(scopes.GetType()))
                {
                    CodeElement elem = point.get_CodeElement(scope);

                    if (elem != null)
                        elems += elem.Name +"  "+elem.StartPoint.Line.ToString()+
                            " (" + scope.ToString() + ")\n";
                }

                MessageBox.Show(
                    "The following elements contain the insertion point:\n\n"
                    + elems);
                CreateToolTipContext(_applicationObject);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void CreateToolTipContext(DTE2 _applicationObject)
        {
            try
            {
                TextSelection selection = (TextSelection)_applicationObject.ActiveDocument.Selection;
                TextPoint point = (TextPoint)selection.ActivePoint;
                // Discover every code element containing the insertion point.
                var classElement = point.get_CodeElement(vsCMElement.vsCMElementClass);
                if (classElement != null)
                {
                    displayLabelOnMethod(classElement, selection);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void displayLabelOnMethod(CodeElement classElement,TextSelection point)
        {
            try
            {
                string className = classElement.FullName.ToString();
                //MethodInfo[] methodInfos = Type.GetType(className).GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
                //foreach(var method in methodInfos)
                //{
                //}
                Label label = AddThisLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private Label AddThisLabel()
        {
            try
            {
                Label label = new Label();
                label.Height = 1000;
                label.Width = 1000;
                label.Location = new Point(100, 100);
                label.Text = "Hello This is label";
                label.Visible = true;
                label.Show();
                label.BringToFront();
                return label;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}