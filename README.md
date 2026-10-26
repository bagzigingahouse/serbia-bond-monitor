# Serbia Bond Monitor

Automatsko praćenje javnih podataka Uprave za javni dug Republike Srbije o **sekundarnom trgovanju državnim kuponskim obveznicama**.

**[Otvori najnoviji pregled trgovanja](docs/latest.md)**  
[Strukturirani podaci (JSON)](docs/latest.json)  
[GitHub Actions — dnevno osvežavanje](../../actions/workflows/update-bonds.yml)

## Kako radi

- GitHub Actions radnim danima u **16:20 UTC** preuzima zvanični fajl [`Sekundarno trgovanje.xlsx`](https://javnidug.gov.rs/static/uploads/Sekundarno%20trgovanje.xlsx). GitHub može zakasniti s početkom zakazanih poslova; dostupnost novih trgovanja zavisi od vremena ažuriranja izvora.
- Program `scripts/build_report.py` koristi samo ugrađene Python biblioteke na GitHubovom serveru. **Nema instalacije Pythona, servera ili lokalnog pokretanja.**
- U izveštaju se za isti datum i ISIN sabiraju promet i broj transakcija, a YTM i cene ponderišu ostvarenim prometom. Izračunava se i promena YTM u baznim poenima u odnosu na prethodno trgovanje *iste serije*.
- Pregled prikazuje poslednjih pet datuma sa ostvarenim trgovanjem, a JSON dodatno poslednjih 45 kalendarskih dana i analizu po preostaloj ročnosti. RSD i EUR se ne sabiraju; restitucione hartije i trezorski zapisi nisu uključeni.
- Ako sajt nije dostupan ili vraća nevažeći fajl, zadatak se označava kao **Failed** i prethodni izveštaj ostaje sačuvan, sa svojim starim datumom. Nema izmišljenih novih podataka.

## Pokretanje i kontrola

Na GitHubu otvori **Actions → Update secondary bond trading → Run workflow**. Prvi rezultat proveravaš u `docs/latest.md` — bitan je datum **poslednje evidentirane transakcije**, ne datum kada je GitHub pokrenuo posao.

Za čitanje iz ChatGPT agenta koristi ovu stalnu adresu:  
`https://github.com/bagzigingahouse/serbia-bond-monitor/blob/main/docs/latest.md`

GitHub Actions podrazumevano koristi dozvolu za pisanje (`contents: write`) definisanu u workflow fajlu. Ako prvi run javi grešku `403` pri `git push`, potrebno je u **Settings → Actions → General → Workflow permissions** uključiti **Read and write permissions**, zatim ponovo pokrenuti workflow.

*Ovo je automatska obrada realizovanih transakcija, ne ponuda za kupovinu/prodaju, niti preporuka za ulaganje.*
