from django import forms

class UploadForm(forms.Form):
    file = forms.FileField(
        label='Select Document (.docx)',
        widget=forms.FileInput(attrs={
            'class': 'absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10', 
            'accept': '.docx'
        })
    )
    url = forms.URLField(
        label='Duty Rates URL',
        initial='https://www.lex.uz/uz/docs/7533457',
        widget=forms.URLInput(attrs={
            'class': 'w-full bg-white border border-gray-300 text-gray-900 placeholder-gray-400 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent block p-3 transition-all duration-200', 
            'placeholder': 'https://...'
        })
    )
