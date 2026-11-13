// Quando o VBA faz dlopen da libprumo.dylib, este construtor inicia o painel (o Swift agenda na thread principal)
// e passa a ouvir "com.gbrein.prumo.show", que o VBA manda com notify_post (botão Painel / carregar o suplemento).
// Chave de segurança: se existir ~/Library/Application Support/Microsoft/Prumo/disable, não faz nada
// (o painel roda dentro do PowerPoint; um defeito nele não pode impedir o PowerPoint de abrir).
#include <dispatch/dispatch.h>
#include <notify.h>
#include <pwd.h>
#include <stdio.h>
#include <unistd.h>

extern void prumo_start(void);
extern void prumo_show(void);
extern void prumo_run_request(void);

__attribute__((constructor)) static void prumo_ctor(void) {
    char path[1024];
    struct passwd *pw = getpwuid(getuid());  // HOME no sandbox aponta para o container; o pw_dir é a casa real
    if (pw && snprintf(path, sizeof path, "%s/Library/Application Support/Microsoft/Prumo/disable", pw->pw_dir) > 0
        && access(path, F_OK) == 0) return;
    int token;
    notify_register_dispatch("com.gbrein.prumo.show", &token, dispatch_get_main_queue(), ^(int t) { (void)t; prumo_show(); });
    // Ação clicada no ribbon: o VBA grava o id em TMPDIR/prumo_run.txt e avisa; o painel pergunta e roda.
    notify_register_dispatch("com.gbrein.prumo.run", &token, dispatch_get_main_queue(), ^(int t) { (void)t; prumo_run_request(); });
    prumo_start();
}
