from pathlib import Path
root=Path(__file__).resolve().parents[2]
for name in ['panel','operator','yonetim']:
    path=root/'templates'/f'{name}.html'
    text=path.read_text(encoding='utf-8').replace('|lower }}',' }}')
    text=text.replace('(arama.value || "").trim().toLocaleLowerCase("tr-TR")','aramaMetni(arama.value)')
    text=text.replace('(kartAra.value || "").trim().toLocaleLowerCase("tr-TR")','aramaMetni(kartAra.value)')
    text=text.replace('adminAra.value\n            .trim()\n            .toLocaleLowerCase("tr-TR")','aramaMetni(adminAra.value)')
    text=text.replace('(kart.dataset.arama || "").includes(metin)','aramaMetni(kart.dataset.arama).includes(metin)')
    text=text.replace('kart.dataset.arama.includes(arama)','aramaMetni(kart.dataset.arama).includes(arama)')
    text=text.replace('satir.dataset.arama.includes(arama)','aramaMetni(satir.dataset.arama).includes(arama)')
    text=text.replace('Dizgiye alınmayı bekleyen ilk 12 kart.','Dizgiye alınmayı bekleyen kartlar; arama ve filtreler tüm kayıtları kapsar.')
    path.write_text(text,encoding='utf-8')
