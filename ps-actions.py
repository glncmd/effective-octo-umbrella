# Opens all the pages of a .pdf file and performs a pre-recorded PS action on each page.
# The script saves each page as a separate .psd file.

# The combination of scripting and PS actions can be very powerful in
# automating most of the repetitive work in Photoshop :)

from comtypes.client import GetActiveObject, CreateObject

app = GetActiveObject("Photoshop.Application", dynamic=True)

start_page = 2
end_page = 45

pdf = "C:\\Users\\Glen\\Desktop\\_.pdf"
dest = "C:\\Users\\Glen\\Desktop\\Result"

psd_options = CreateObject("Photoshop.PhotoshopSaveOptions", dynamic=True)
pdf_options = CreateObject("Photoshop.PDFOpenOptions", dynamic=True)

def play_action(action_name: str):
    '''
    Performs Photoshop Action
    '''
    id_ply = app.CharIDToTypeID("Ply ")
    desc = CreateObject("Photoshop.ActionDescriptor", dynamic=True)
    id_null = app.CharIDToTypeID("null")
    ref = CreateObject("Photoshop.ActionReference", dynamic=True)
    id_actn = app.CharIDToTypeID("Actn")
    ref.PutName(id_actn, action_name)
    id_action_set = app.CharIDToTypeID("ASet")
    ref.PutName(id_action_set, "Default Actions")
    desc.PutReference(id_null, ref)
    app.ExecuteAction(id_ply, desc, 3)


for i in range(start_page, end_page + 1):
    print(f'Current Page: {i}/{end_page}')

    pdf_options.Page = i
    app.Open(pdf, pdf_options, True)

    play_action('Sample_Action')

    current_doc = app.ActiveDocument
    current_doc.ActiveLayer = current_doc.Layers.Item(1)
    current_doc.SaveAs(f'{dest}\\{i}', psd_options)
    current_doc.Close()
