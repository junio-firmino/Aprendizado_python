from win10toast import ToastNotifier

toast = ToastNotifier()

toast.show_toast(title='Notificação de teste',
                 msg='Teste realizado com sucesso',
                 duration=10,
                 icon_path=None)